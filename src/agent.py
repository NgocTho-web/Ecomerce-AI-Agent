"""
AI Agent phân tích dữ liệu bán hàng cho Ecommerce-AI-Agent.

Yêu cầu:
    pip install langchain-core langchain-ollama

Cấu hình:
    - Ollama đang chạy
    - Đã tải model: ollama pull qwen2.5:7b
    - src/tools/data_tools.py có các tool:
      query_filter_data, calculate_metrics, generate_chart, analyze_products
"""

from __future__ import annotations

import ast
import json
import re
from typing import Any, Optional

from langchain_core.messages import AIMessage, HumanMessage, SystemMessage, ToolMessage
from langchain_ollama import ChatOllama

from src.tools.data_tools import (
    analyze_products,
    calculate_metrics,
    generate_chart,
    query_filter_data,
)

# ---------------------------------------------------------------------
# 1. Khai báo công cụ
# ---------------------------------------------------------------------

TOOLS = [
    query_filter_data,
    calculate_metrics,
    generate_chart,
    analyze_products,
]
TOOL_MAP = {tool.name: tool for tool in TOOLS}

MODEL_NAME = "qwen2.5:7b"
TEMPERATURE = 0
MAX_ITERATIONS = 8

SYSTEM_PROMPT = """
Bạn là AI Agent chuyên phân tích dữ liệu bán hàng và hiệu quả kinh doanh
từ bộ dữ liệu Superstore.

Quy tắc làm việc:
1. Dùng công cụ khi cần số liệu. Không tự bịa số liệu, không ước lượng số liệu
   rồi trình bày như dữ liệu thật.
2. Dùng calculate_metrics để tính KPI/tổng hợp; query_filter_data để truy vấn
   các bản ghi; generate_chart để tạo biểu đồ; analyze_products để xếp hạng
   sản phẩm theo doanh số/lợi nhuận/số lượng.
3. Đọc kỹ kết quả tool trước khi kết luận. Nếu tool trả về lỗi hoặc không có dữ
   liệu, hãy nói rõ thay vì suy diễn.
4. Phân biệt Sales (doanh số) và Profit (lợi nhuận). Doanh số cao không đồng
   nghĩa với lợi nhuận cao.
5. Khi nhận xét sản phẩm bán chậm, phải nói rõ đây là đánh giá tương đối theo
   chỉ số và phạm vi dữ liệu hiện có. Nếu không có dữ liệu tồn kho hoặc thời
   gian đủ chi tiết, không được khẳng định sản phẩm thực sự tồn kho lâu.
6. Trả lời bằng tiếng Việt, rõ ràng, có cấu trúc. Khi phù hợp, nêu số liệu,
   đơn vị, nhóm so sánh và ý nghĩa kinh doanh.
7. Không tiết lộ tên khách hàng hay thông tin cá nhân trong dữ liệu.
8. Nếu người dùng yêu cầu biểu đồ, hãy gọi generate_chart và giữ nguyên
   chart_path/chart_filename do tool trả về. Không tự bịa đường dẫn.
9. Khi người dùng yêu cầu phân tích "theo danh mục",
   "theo khu vực", "theo phân khúc" hoặc "theo thời gian",
   hãy hiểu đó là yêu cầu GROUP BY trên toàn bộ các nhóm,
   không phải lọc riêng một nhóm.

10. Tham số category, segment, region và year của
    calculate_metrics là bộ lọc tùy chọn. Nếu người dùng
    không yêu cầu lọc cụ thể, hãy bỏ qua các tham số đó.

11. Nếu người dùng yêu cầu nhiều chỉ số, ví dụ Sales và
    Profit, hãy gọi calculate_metrics riêng cho từng chỉ số.
    Nếu muốn so sánh theo danh mục, hãy đặt group_by="Category".

12. Không yêu cầu người dùng bổ sung thông tin nếu yêu cầu
    đã đủ rõ để thực hiện trên toàn bộ dữ liệu.

"""


def _make_llm(model: str = MODEL_NAME, temperature: float = TEMPERATURE):
    """Khởi tạo Ollama và bật tool calling."""
    llm = ChatOllama(model=model, temperature=temperature)
    return llm.bind_tools(TOOLS)


def _to_jsonable(value: Any) -> Any:
    """Chuẩn hóa kết quả tool để đưa vào ToolMessage/JSON."""
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    if isinstance(value, dict):
        return {str(k): _to_jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_to_jsonable(v) for v in value]
    # Hỗ trợ các kiểu scalar của pandas/numpy mà không phụ thuộc trực tiếp numpy.
    if hasattr(value, "item"):
        try:
            return _to_jsonable(value.item())
        except Exception:
            pass
    return str(value)


def _tool_content(value: Any) -> str:
    """Định dạng kết quả tool thành chuỗi ổn định cho LLM."""
    value = _to_jsonable(value)
    try:
        return json.dumps(value, ensure_ascii=False, default=str)
    except Exception:
        return str(value)


def _invoke_tool(tool_name: str, tool_args: dict[str, Any]) -> Any:
    """Gọi tool theo tên và bắt lỗi để agent có thể báo cáo rõ ràng."""
    tool = TOOL_MAP.get(tool_name)
    if tool is None:
        return {"status": "error", "error": f"Không tìm thấy công cụ: {tool_name}"}

    try:
        result = tool.invoke(tool_args)
        # Một số phiên bản LangChain có thể trả ToolMessage/ToolException.
        if isinstance(result, ToolMessage):
            return {"status": "success", "result": result.content}
        return _to_jsonable(result)
    except Exception as exc:
        return {
            "status": "error",
            "tool": tool_name,
            "error": f"{type(exc).__name__}: {exc}",
        }


def _collect_artifacts(value: Any, artifacts: list[dict[str, str]]) -> None:
    """Tìm đường dẫn biểu đồ trong kết quả tool, không dựa vào văn bản LLM."""
    if isinstance(value, dict):
        path = value.get("chart_path")
        if isinstance(path, str) and path:
            item = {
                "chart_path": path,
                "chart_filename": str(value.get("chart_filename") or ""),
            }
            if item not in artifacts:
                artifacts.append(item)
        for child in value.values():
            _collect_artifacts(child, artifacts)
    elif isinstance(value, list):
        for child in value:
            _collect_artifacts(child, artifacts)


def _parse_tool_args(raw_args: Any) -> dict[str, Any]:
    """Chuẩn hóa arguments do model trả về."""
    if isinstance(raw_args, dict):
        return raw_args
    if isinstance(raw_args, str):
        try:
            parsed = json.loads(raw_args)
            if isinstance(parsed, dict):
                return parsed
        except json.JSONDecodeError:
            pass
        try:
            parsed = ast.literal_eval(raw_args)
            if isinstance(parsed, dict):
                return parsed
        except (ValueError, SyntaxError):
            pass
    return {}


def run_agent_loop(
    user_input: str,
    max_iterations: int = MAX_ITERATIONS,
    model: str = MODEL_NAME,
) -> dict[str, Any]:
    """
    Chạy vòng lặp suy luận/tool calling.

    Trả về:
        {
            "success": bool,
            "answer": str,
            "chart_path": str | None,
            "charts": list[dict],
            "tool_results": list[dict],
            "iterations": int
        }
    """
    if not isinstance(user_input, str) or not user_input.strip():
        return {
            "success": False,
            "answer": "Vui lòng cung cấp một yêu cầu phân tích.",
            "chart_path": None,
            "charts": [],
            "tool_results": [],
            "iterations": 0,
        }

    try:
        llm_with_tools = _make_llm(model=model)
    except Exception as exc:
        return {
            "success": False,
            "answer": f"Không khởi tạo được Ollama: {type(exc).__name__}: {exc}",
            "chart_path": None,
            "charts": [],
            "tool_results": [],
            "iterations": 0,
        }

    messages = [
        SystemMessage(content=SYSTEM_PROMPT),
        HumanMessage(content=user_input.strip()),
    ]
    tool_results: list[dict[str, Any]] = []
    charts: list[dict[str, str]] = []

    for iteration in range(1, max_iterations + 1):
        try:
            response: AIMessage = llm_with_tools.invoke(messages)
        except Exception as exc:
            return {
                "success": False,
                "answer": (
                    "Không gọi được mô hình Ollama. Hãy kiểm tra Ollama đang chạy, "
                    f"model '{model}' đã được tải và tương thích với tool calling.\n"
                    f"Chi tiết: {type(exc).__name__}: {exc}"
                ),
                "chart_path": charts[-1]["chart_path"] if charts else None,
                "charts": charts,
                "tool_results": tool_results,
                "iterations": iteration,
            }

        messages.append(response)
        calls = getattr(response, "tool_calls", None) or []

        # Không còn yêu cầu gọi tool: trả câu trả lời cuối của LLM.
        if not calls:
            answer = response.content
            if isinstance(answer, list):
                answer = "\n".join(
                    str(part.get("text", "")) if isinstance(part, dict) else str(part)
                    for part in answer
                )
            answer = str(answer).strip() or "Đã hoàn tất xử lý nhưng mô hình không trả lời dạng văn bản."
            return {
                "success": not any(r.get("status") == "error" for r in tool_results),
                "answer": answer,
                "chart_path": charts[-1]["chart_path"] if charts else None,
                "charts": charts,
                "tool_results": tool_results,
                "iterations": iteration,
            }

        for call in calls:
            tool_name = call.get("name", "")
            tool_args = _parse_tool_args(call.get("args", {}))
            tool_call_id = call.get("id") or f"{tool_name}_{iteration}"

            result = _invoke_tool(tool_name, tool_args)
            _collect_artifacts(result, charts)
            tool_results.append({
                "tool": tool_name,
                "arguments": tool_args,
                "result": result,
            })
            messages.append(
                ToolMessage(
                    content=_tool_content(result),
                    tool_call_id=tool_call_id,
                )
            )

    # Nếu hết lượt, yêu cầu mô hình tóm tắt những gì đã thu được mà không gọi tool.
    try:
        final_messages = messages + [
            HumanMessage(
                content=(
                    "Hãy tổng hợp câu trả lời cuối cùng bằng tiếng Việt chỉ từ kết quả "
                    "công cụ đã có ở trên. Không gọi thêm công cụ và không bịa số liệu."
                )
            )
        ]
        final_llm = ChatOllama(model=model, temperature=TEMPERATURE)
        final_response = final_llm.invoke(final_messages)
        final_answer = str(final_response.content).strip()
    except Exception:
        final_answer = (
            "Đã chạy đến giới hạn số vòng lặp. Hãy xem tool_results để kiểm tra "
            "các kết quả công cụ đã thực thi."
        )

    return {
        "success": not any(r.get("status") == "error" for r in tool_results),
        "answer": final_answer,
        "chart_path": charts[-1]["chart_path"] if charts else None,
        "charts": charts,
        "tool_results": tool_results,
        "iterations": max_iterations,
    }


def ask_agent(query: str) -> tuple[bool, str, Optional[str]]:
    """
    Hàm tương thích với cách gọi đơn giản ở dashboard cũ.

    Returns:
        (success, answer, chart_path)
    """
    result = run_agent_loop(query)
    return (
        bool(result.get("success")),
        str(result.get("answer", "")),
        result.get("chart_path"),
    )


# ---------------------------------------------------------------------
# 2. Chế độ phân tích tự động
# ---------------------------------------------------------------------

DEFAULT_ANALYSIS_TASKS = [
    {
        "task": "Tổng doanh số, lợi nhuận và số lượng bán ra",
        "query": (
            "Dùng công cụ calculate_metrics để tính tổng Sales, tổng Profit và "
            "tổng Quantity trên toàn bộ dữ liệu. Trình bày rõ từng chỉ số."
        ),
    },
    {
        "task": "Hiệu quả kinh doanh theo danh mục",
        "query": (
            "Dùng calculate_metrics để tổng hợp Sales và Profit theo Category. "
            "So sánh doanh số với lợi nhuận và chỉ ra danh mục cần chú ý."
        ),
    },
    {
        "task": "Nhóm sản phẩm có doanh số thấp",
        "query": (
            "Dùng analyze_products để xếp hạng các sản phẩm theo Sales tăng dần, "
            "lấy 10 sản phẩm cuối bảng theo doanh số thấp nhất. Nêu giới hạn của "
            "việc dùng doanh số để suy ra bán chậm."
        ),
    },
    {
        "task": "Xu hướng doanh số theo thời gian",
        "query": (
            "Dùng calculate_metrics để tổng hợp Sales theo YearMonth và dùng "
            "generate_chart để tạo biểu đồ line với YearMonth trên trục X, Sales "
            "trên trục Y, aggregation sum. Tóm tắt xu hướng dựa trên kết quả."
        ),
    },
]


def run_automatic_analysis(
    tasks: Optional[list[dict[str, str]]] = None,
    model: str = MODEL_NAME,
) -> dict[str, Any]:
    """
    Chạy lần lượt một bộ nhiệm vụ đã định nghĩa, không cần người dùng hỏi từng câu.

    Có thể truyền tasks tùy chỉnh dạng:
        [{"task": "Tên nhiệm vụ", "query": "Yêu cầu phân tích"}]
    """
    selected_tasks = tasks if tasks is not None else DEFAULT_ANALYSIS_TASKS
    reports: list[dict[str, Any]] = []
    all_charts: list[dict[str, str]] = []

    for task in selected_tasks:
        task_name = str(task.get("task", "Nhiệm vụ phân tích"))
        query = str(task.get("query", "")).strip()
        if not query:
            reports.append({
                "task": task_name,
                "success": False,
                "answer": "Bỏ qua nhiệm vụ vì query đang trống.",
                "chart_path": None,
            })
            continue

        result = run_agent_loop(query, model=model)
        for chart in result.get("charts", []):
            if chart not in all_charts:
                all_charts.append(chart)

        reports.append({
            "task": task_name,
            "success": result.get("success", False),
            "answer": result.get("answer", ""),
            "chart_path": result.get("chart_path"),
            "charts": result.get("charts", []),
            "tool_results": result.get("tool_results", []),
        })

    return {
        "success": all(report["success"] for report in reports) if reports else False,
        "reports": reports,
        "charts": all_charts,
        "summary": (
            f"Đã hoàn thành {sum(1 for r in reports if r['success'])}/"
            f"{len(reports)} nhiệm vụ phân tích."
        ),
    }


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Ecommerce AI Agent")
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Chạy bộ phân tích tự động định sẵn.",
    )
    parser.add_argument(
        "--query",
        type=str,
        default=None,
        help="Chạy một yêu cầu phân tích đơn lẻ.",
    )
    parser.add_argument("--model", type=str, default=MODEL_NAME)
    args = parser.parse_args()

    if args.auto:
        output = run_automatic_analysis(model=args.model)
        print(output["summary"])
        for report in output["reports"]:
            print(f"\n{'=' * 70}\n{report['task']}\n{'=' * 70}")
            print(report["answer"])
            if report.get("chart_path"):
                print(f"Biểu đồ: {report['chart_path']}")
    else:
        query = args.query or input("Nhập yêu cầu phân tích: ").strip()
        output = run_agent_loop(query, model=args.model)
        print(output["answer"])
        if output.get("chart_path"):
            print(f"\nBiểu đồ: {output['chart_path']}")

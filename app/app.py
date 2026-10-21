"""
Giao diện Streamlit cho Ecommerce AI Agent (Thành viên 4).

Chạy (từ thư mục gốc dự án):
    streamlit run app/app.py

Điều kiện: Ollama đang chạy và đã tải model (mặc định xem src/agent.py -> MODEL_NAME).
"""
from __future__ import annotations

import csv
import json
import os
import sys
import time
from datetime import datetime

import pandas as pd
import streamlit as st

# Cho phép `from src...` khi chạy bằng `streamlit run app/app.py`
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

st.set_page_config(page_title="Ecommerce AI Agent", page_icon="📊", layout="wide")

LOG_PATH = os.path.join(BASE_DIR, "logs", "usage_log.csv")
MAX_QUERY_CHARS = 1000

EXAMPLES = {
    "Thống kê": [
        "Tổng doanh số và tổng lợi nhuận là bao nhiêu?",
        "Doanh số trung bình trên mỗi dòng đơn hàng của phân khúc Consumer?",
    ],
    "Lọc dữ liệu": [
        "Có bao nhiêu dòng đơn hàng bị lỗ ở danh mục Furniture?",
        "Cho xem 5 đơn hàng Phones ở California.",
    ],
    "So sánh": [
        "So sánh doanh số và lợi nhuận giữa các danh mục.",
        "Khu vực nào có lợi nhuận cao nhất?",
    ],
    "Xu hướng & biểu đồ": [
        "Vẽ biểu đồ đường doanh số theo tháng và nhận xét xu hướng.",
        "Vẽ biểu đồ cột lợi nhuận theo Sub-Category.",
    ],
    "Sản phẩm": [
        "Top 5 sản phẩm bị lỗ nặng nhất.",
        "10 sản phẩm có doanh số thấp nhất.",
    ],
}


# ---------------------------------------------------------------------------
# Nạp agent (báo lỗi thân thiện thay vì làm sập giao diện)
# ---------------------------------------------------------------------------
@st.cache_resource(show_spinner="Đang khởi tạo agent và nạp dữ liệu...")
def load_backend():
    from src import agent as agent_module
    from src.tools import data_tools
    return agent_module, data_tools


try:
    agent_module, data_tools = load_backend()
    BACKEND_ERROR = None
except Exception as exc:  # thiếu thư viện, thiếu dữ liệu, ...
    agent_module, data_tools = None, None
    BACKEND_ERROR = f"{type(exc).__name__}: {exc}"


# ---------------------------------------------------------------------------
# Tiện ích
# ---------------------------------------------------------------------------
def log_usage(query: str, result: dict, seconds: float, model: str) -> None:
    """Ghi nhật ký mỗi lượt hỏi để Thành viên 4 thống kê thời gian phản hồi / tỷ lệ lỗi."""
    try:
        os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
        new = not os.path.isfile(LOG_PATH)
        with open(LOG_PATH, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["time", "model", "query", "success", "seconds", "iterations", "tools", "has_chart"])
            w.writerow([
                datetime.now().isoformat(timespec="seconds"), model, query, result.get("success"),
                round(seconds, 2), result.get("iterations"),
                ",".join(t.get("tool", "") for t in result.get("tool_results", [])),
                bool(result.get("charts")),
            ])
    except OSError:
        pass  # không để lỗi ghi log ảnh hưởng người dùng


def tables_from_tool_results(tool_results: list[dict]) -> list[tuple[str, pd.DataFrame]]:
    """Lấy các bảng số liệu từ kết quả tool để hiển thị (không dựa vào văn bản LLM)."""
    tables: list[tuple[str, pd.DataFrame]] = []
    for item in tool_results:
        res = item.get("result")
        if not isinstance(res, dict):
            continue
        for key in ("results", "sample_records", "products"):
            rows = res.get(key)
            if isinstance(rows, list) and rows and isinstance(rows[0], dict):
                label = f"{item.get('tool')} → {key}"
                tables.append((label, pd.DataFrame(rows)))
    return tables


def render_result(msg: dict, key_prefix: str) -> None:
    """Hiển thị một câu trả lời của agent: văn bản, biểu đồ, bảng, chi tiết thực thi."""
    if msg.get("success") is False:
        st.warning("Agent báo có lỗi trong quá trình xử lý – xem chi tiết bên dưới.")
    st.markdown(msg["answer"])

    for i, chart in enumerate(msg.get("charts", [])):
        path = chart.get("chart_path", "")
        if path and os.path.isfile(path):
            st.image(path, caption=chart.get("chart_filename") or os.path.basename(path), use_container_width=True)
        elif path:
            st.caption(f"Không tìm thấy file biểu đồ: {path}")

    tables = tables_from_tool_results(msg.get("tool_results", []))
    if tables:
        with st.expander(f"Bảng số liệu từ công cụ ({len(tables)})"):
            for label, df in tables:
                st.caption(label)
                st.dataframe(df, use_container_width=True, hide_index=True)

    meta = f"⏱ {msg.get('seconds', 0):.1f}s · 🔁 {msg.get('iterations', 0)} vòng · 🧰 {len(msg.get('tool_results', []))} lần gọi công cụ"
    with st.expander(f"Chi tiết thực thi — {meta}"):
        if not msg.get("tool_results"):
            st.write("Agent không gọi công cụ nào.")
        for n, item in enumerate(msg.get("tool_results", []), 1):
            st.markdown(f"**{n}. `{item.get('tool')}`**")
            st.code(json.dumps(item.get("arguments", {}), ensure_ascii=False, indent=2), language="json")
            res = item.get("result")
            text = json.dumps(res, ensure_ascii=False, indent=2, default=str)
            st.code(text if len(text) < 4000 else text[:4000] + "\n... (đã cắt bớt)", language="json")


def ask(query: str, model: str) -> dict:
    t0 = time.perf_counter()
    result = agent_module.run_agent_loop(query, model=model)
    seconds = time.perf_counter() - t0
    msg = {
        "role": "assistant",
        "answer": result.get("answer", ""),
        "success": result.get("success"),
        "charts": result.get("charts", []),
        "tool_results": result.get("tool_results", []),
        "iterations": result.get("iterations", 0),
        "seconds": seconds,
    }
    log_usage(query, result, seconds, model)
    return msg


# ---------------------------------------------------------------------------
# Giao diện
# ---------------------------------------------------------------------------
st.title("📊 Ecommerce AI Agent")
st.caption("Hỏi đáp và phân tích dữ liệu bán hàng Superstore bằng ngôn ngữ tự nhiên.")

if BACKEND_ERROR:
    st.error("Không khởi tạo được AI Agent.")
    st.code(BACKEND_ERROR)
    st.info("Kiểm tra: `pip install -r requirements.txt`, file dữ liệu "
            "`data/processed/ecommerce_cleaned_for_agent.csv`, và chạy lệnh từ thư mục gốc dự án.")
    st.stop()

if "history" not in st.session_state:
    st.session_state.history = []          # [{"role": "user"/"assistant", ...}]
if "pending" not in st.session_state:
    st.session_state.pending = None        # câu hỏi mẫu được bấm ở sidebar

df_all = data_tools.df_global

with st.sidebar:
    st.header("Cài đặt")
    model = st.text_input("Model Ollama", value=agent_module.MODEL_NAME,
                          help="Phải đã được tải bằng `ollama pull <tên model>`.")
    if st.button("🗑 Xóa lịch sử chat", use_container_width=True):
        st.session_state.history = []
        st.rerun()

    st.divider()
    st.header("Câu hỏi gợi ý")
    for group, qs in EXAMPLES.items():
        with st.expander(group):
            for j, q in enumerate(qs):
                if st.button(q, key=f"ex_{group}_{j}", use_container_width=True):
                    st.session_state.pending = q

    st.divider()
    valid_dates = df_all["Order Date"].notna().mean() if "Order Date" in df_all.columns else 0
    st.caption(f"Dữ liệu: {len(df_all):,} dòng · {df_all.shape[1]} cột")
    if valid_dates < 0.95:
        st.warning(f"Chỉ {valid_dates:.0%} dòng có Order Date hợp lệ – kết quả theo năm/tháng có thể thiếu.")

tab_chat, tab_auto, tab_data = st.tabs(["💬 Hỏi đáp", "⚙️ Phân tích tự động", "🗂 Dữ liệu"])

# ---- Tab 1: Chat ----
with tab_chat:
    for i, msg in enumerate(st.session_state.history):
        with st.chat_message(msg["role"]):
            if msg["role"] == "user":
                st.markdown(msg["content"])
            else:
                render_result(msg, key_prefix=f"h{i}")

    user_query = st.chat_input("Nhập câu hỏi phân tích, ví dụ: Danh mục nào có lợi nhuận cao nhất?")
    if st.session_state.pending:
        user_query, st.session_state.pending = st.session_state.pending, None

    if user_query is not None:
        text = user_query.strip()
        if not text:
            st.warning("Vui lòng nhập câu hỏi.")
        elif len(text) > MAX_QUERY_CHARS:
            st.warning(f"Câu hỏi quá dài (tối đa {MAX_QUERY_CHARS} ký tự).")
        else:
            st.session_state.history.append({"role": "user", "content": text})
            with st.chat_message("user"):
                st.markdown(text)
            with st.chat_message("assistant"):
                with st.spinner("Agent đang phân tích dữ liệu..."):
                    reply = ask(text, model)
                render_result(reply, key_prefix="new")
            st.session_state.history.append(reply)

# ---- Tab 2: Phân tích tự động ----
with tab_auto:
    st.write("Chạy bộ phân tích định sẵn (KPI tổng, theo danh mục, sản phẩm doanh số thấp, xu hướng theo tháng).")
    if st.button("▶ Chạy phân tích tự động"):
        tasks = agent_module.DEFAULT_ANALYSIS_TASKS
        progress = st.progress(0.0, text="Bắt đầu...")
        reports = []
        for k, task in enumerate(tasks):
            progress.progress(k / len(tasks), text=f"Đang chạy: {task['task']} ({k + 1}/{len(tasks)})")
            res = agent_module.run_agent_loop(task["query"], model=model)
            reports.append({"task": task["task"], **res})
        progress.progress(1.0, text="Hoàn tất")
        st.session_state.auto_reports = reports
    for rep in st.session_state.get("auto_reports", []):
        with st.container(border=True):
            st.subheader(rep["task"])
            render_result({
                "answer": rep.get("answer", ""), "success": rep.get("success"),
                "charts": rep.get("charts", []), "tool_results": rep.get("tool_results", []),
                "iterations": rep.get("iterations", 0), "seconds": 0,
            }, key_prefix=rep["task"])

# ---- Tab 3: Dữ liệu ----
with tab_data:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng doanh số", f"${df_all['Sales'].sum():,.0f}")
    c2.metric("Tổng lợi nhuận", f"${df_all['Profit'].sum():,.0f}")
    c3.metric("Tổng số lượng", f"{int(df_all['Quantity'].sum()):,}")
    c4.metric("Số đơn hàng", f"{df_all['Order ID'].nunique():,}" if "Order ID" in df_all.columns else "n/a")
    st.caption("Các số trên tính trực tiếp từ dữ liệu, dùng để đối chiếu câu trả lời của agent.")
    shown = [c for c in df_all.columns if c not in ("Customer ID",)]  # không hiển thị định danh khách hàng
    st.dataframe(df_all[shown].head(200), use_container_width=True, hide_index=True)

import os
import uuid
from typing import Optional

import pandas as pd
import matplotlib
matplotlib.use("Agg")  # Phù hợp chạy pipeline/server không có màn hình
import matplotlib.pyplot as plt
import seaborn as sns
from langchain_core.tools import tool

from src.config import CLEANED_DATA_PATH, CHART_OUTPUT_DIR
from src.schemas import FilterSalesInput, CalculateMetricsInput, GenerateChartInput


# =============================================================================
# Cấu hình và nạp dữ liệu
# =============================================================================

NUMERIC_COLUMNS = {"Sales", "Profit", "Quantity", "Discount"}
FILTER_COLUMNS = {
    "Category", "Sub-Category", "Segment", "Region", "State",
    "Ship Mode", "Year", "YearMonth",
}
AGGREGATIONS = {"sum", "mean", "median", "min", "max", "count"}
CHART_TYPES = {"bar", "line", "pie"}

sns.set_theme(style="whitegrid")


def _load_dataframe() -> pd.DataFrame:
    """Đọc dữ liệu sạch và chuẩn hóa các cột ngày/thời gian cần thiết."""
    path = os.fspath(CLEANED_DATA_PATH)

    if not os.path.isfile(path):
        raise FileNotFoundError(
            f"Không tìm thấy dữ liệu sạch tại: {path}. "
            "Hãy kiểm tra CLEANED_DATA_PATH hoặc chạy notebook làm sạch dữ liệu."
        )

    data = pd.read_csv(path)

    required = {"Order Date", "Sales", "Profit", "Quantity"}
    missing = sorted(required - set(data.columns))
    if missing:
        raise ValueError(
            f"Dữ liệu thiếu các cột bắt buộc: {missing}. "
            f"Các cột hiện có: {list(data.columns)}"
        )

    for date_col in ("Order Date", "Ship Date"):
        if date_col in data.columns:
            data[date_col] = pd.to_datetime(data[date_col], errors="coerce")

    data["Sales"] = pd.to_numeric(data["Sales"], errors="coerce")
    data["Profit"] = pd.to_numeric(data["Profit"], errors="coerce")
    data["Quantity"] = pd.to_numeric(data["Quantity"], errors="coerce")
    if "Discount" in data.columns:
        data["Discount"] = pd.to_numeric(data["Discount"], errors="coerce")

    data["Year"] = data["Order Date"].dt.year
    data["YearMonth"] = data["Order Date"].dt.to_period("M").astype("string")

    return data


# Cache dữ liệu trong bộ nhớ để các tool cùng dùng một bản dữ liệu.
# Gọi refresh_data() khi pipeline phát hiện dữ liệu mới.
df_global = _load_dataframe()


def refresh_data() -> dict:
    """Nạp lại dữ liệu từ CLEANED_DATA_PATH; gọi khi dữ liệu đầu vào được cập nhật."""
    global df_global
    df_global = _load_dataframe()
    return {
        "status": "success",
        "rows": int(len(df_global)),
        "columns": list(df_global.columns),
        "source": os.fspath(CLEANED_DATA_PATH),
    }


def _get_data() -> pd.DataFrame:
    """Trả về bản sao dữ liệu để tránh tool vô tình sửa dữ liệu dùng chung."""
    return df_global.copy()


def _filter_data(
    data: pd.DataFrame,
    category: Optional[str] = None,
    sub_category: Optional[str] = None,
    segment: Optional[str] = None,
    region: Optional[str] = None,
    state: Optional[str] = None,
    year: Optional[int] = None,
    is_loss: Optional[bool] = None,
) -> pd.DataFrame:
    """Lọc dữ liệu theo các điều kiện chung."""
    conditions = [
        ("Category", category),
        ("Sub-Category", sub_category),
        ("Segment", segment),
        ("Region", region),
        ("State", state),
    ]

    for column, value in conditions:
        if value is not None:
            if column not in data.columns:
                return data.iloc[0:0]
            data = data[
                data[column].astype("string").str.casefold()
                == str(value).casefold()
            ]

    if year is not None:
        if "Year" not in data.columns:
            return data.iloc[0:0]
        data = data[data["Year"] == year]

    if is_loss is not None:
        if "Profit" not in data.columns:
            return data.iloc[0:0]
        data = data[data["Profit"] < 0] if is_loss else data[data["Profit"] >= 0]

    return data


def _json_safe_records(data: pd.DataFrame) -> list[dict]:
    """Chuyển bản ghi thành kiểu dữ liệu có thể serialize thành JSON."""
    safe = data.astype(object).where(pd.notna(data), None)
    records = safe.to_dict(orient="records")
    for record in records:
        for key, value in record.items():
            if isinstance(value, pd.Timestamp):
                record[key] = value.isoformat()
    return records


# =============================================================================
# Tool 1: Lọc và truy vấn dữ liệu
# =============================================================================

@tool(args_schema=FilterSalesInput)
def query_filter_data(
    category: str = None,
    sub_category: str = None,
    segment: str = None,
    region: str = None,
    state: str = None,
    year: int = None,
    is_loss: bool = None,
    limit: int = 5,
) -> dict:
    """Lọc dữ liệu Superstore theo danh mục, phân khúc, vùng, bang, năm hoặc trạng thái lãi/lỗ; trả tổng số dòng khớp và một mẫu bản ghi."""
    try:
        if limit < 1 or limit > 100:
            return {"error": "limit phải nằm trong khoảng 1 đến 100."}

        data = _filter_data(
            _get_data(),
            category=category,
            sub_category=sub_category,
            segment=segment,
            region=region,
            state=state,
            year=year,
            is_loss=is_loss,
        )

        if data.empty:
            return {
                "total_matched": 0,
                "sample_records": [],
                "message": "Không tìm thấy bản ghi nào khớp với điều kiện lọc.",
            }

        display_columns = [
            column for column in (
                "Order ID", "Order Date", "Segment", "Category",
                "Sub-Category", "Region", "State", "Sales", "Profit", "Quantity"
            )
            if column in data.columns
        ]

        return {
            "total_matched": int(len(data)),
            "sample_records": _json_safe_records(data[display_columns].head(limit)),
        }
    except Exception as exc:
        return {"error": f"Không thể lọc dữ liệu: {exc}"}


# =============================================================================
# Tool 2: Tính toán KPI/thống kê
# =============================================================================

@tool(args_schema=CalculateMetricsInput)
def calculate_metrics(
    metric: str,
    aggregation: str = "sum",
    group_by: str | None = None,
    category: str | None = None,
    segment: str | None = None,
    region: str | None = None,
    year: int | None = None,
) -> dict:
    """Tính tổng, trung bình, trung vị, nhỏ nhất, lớn nhất hoặc số giá trị cho Sales, Profit, Quantity, Discount; có thể phân nhóm theo cột được hỗ trợ."""
    try:
        if metric not in NUMERIC_COLUMNS:
            return {
                "error": f"metric không hợp lệ: {metric}. Chọn một trong {sorted(NUMERIC_COLUMNS)}."
            }
        if aggregation not in AGGREGATIONS:
            return {
                "error": f"aggregation không hợp lệ: {aggregation}. Chọn một trong {sorted(AGGREGATIONS)}."
            }
        if group_by is not None and group_by not in FILTER_COLUMNS:
            return {
                "error": f"group_by không hợp lệ: {group_by}. Chọn một trong {sorted(FILTER_COLUMNS)}."
            }

        data = _filter_data(
            _get_data(),
            category=category,
            segment=segment,
            region=region,
            year=year,
        )

        if metric not in data.columns:
            return {"error": f"Dữ liệu không có cột {metric}."}
        if data.empty:
            return {"error": "Không có dữ liệu thỏa mãn điều kiện lọc."}

        data[metric] = pd.to_numeric(data[metric], errors="coerce")
        data = data.dropna(subset=[metric])
        if data.empty:
            return {"error": f"Cột {metric} không có giá trị số hợp lệ."}

        if group_by:
            if group_by not in data.columns:
                return {"error": f"Dữ liệu không có cột phân nhóm {group_by}."}
            grouped = (
                data.groupby(group_by, dropna=False)[metric]
                .agg(aggregation)
                .round(2)
                .reset_index()
            )
            return {
                "metric": metric,
                "aggregation": aggregation,
                "group_by": group_by,
                "row_count": int(len(data)),
                "results": _json_safe_records(grouped),
            }

        value = getattr(data[metric], aggregation)()
        return {
            "metric": metric,
            "aggregation": aggregation,
            "value": None if pd.isna(value) else round(float(value), 2),
            "row_count": int(len(data)),
            "unique_orders": (
                int(data["Order ID"].nunique())
                if "Order ID" in data.columns else None
            ),
        }
    except Exception as exc:
        return {"error": f"Không thể tính chỉ số: {exc}"}


# =============================================================================
# Tool 3: Tạo biểu đồ
# =============================================================================

@tool(args_schema=GenerateChartInput)
def generate_chart(
    chart_type: str,
    x_axis: str,
    y_axis: str = "Sales",
    aggregation: str = "sum",
    chart_title: str = "Biểu đồ phân tích kinh doanh",
) -> dict:
    """Tạo và lưu biểu đồ bar, line hoặc pie từ dữ liệu bán hàng; trả tên file và đường dẫn ảnh để dashboard sử dụng."""
    try:
        if chart_type not in CHART_TYPES:
            return {"error": f"chart_type phải là một trong {sorted(CHART_TYPES)}."}
        if y_axis not in {"Sales", "Profit", "Quantity"}:
            return {"error": "y_axis phải là Sales, Profit hoặc Quantity."}
        if aggregation not in {"sum", "mean"}:
            return {"error": "aggregation phải là sum hoặc mean."}
        if x_axis not in FILTER_COLUMNS:
            return {"error": f"x_axis không hợp lệ: {x_axis}."}

        data = _get_data()
        if x_axis not in data.columns or y_axis not in data.columns:
            return {"error": f"Không tìm thấy cột {x_axis} hoặc {y_axis} trong dữ liệu."}

        data[y_axis] = pd.to_numeric(data[y_axis], errors="coerce")
        data = data.dropna(subset=[x_axis, y_axis])
        if data.empty:
            return {"error": "Không có dữ liệu hợp lệ để vẽ biểu đồ."}

        grouped = (
            data.groupby(x_axis, dropna=False)[y_axis]
            .agg(aggregation)
            .reset_index()
        )
        if grouped.empty:
            return {"error": "Không có dữ liệu sau khi tổng hợp để vẽ biểu đồ."}

        # Giới hạn số nhóm để biểu đồ dễ đọc; biểu đồ đường ưu tiên thứ tự thời gian.
        if len(grouped) > 30:
            grouped = grouped.nlargest(30, y_axis)

        os.makedirs(CHART_OUTPUT_DIR, exist_ok=True)
        filename = f"chart_{uuid.uuid4().hex[:10]}.png"
        save_path = os.path.join(os.fspath(CHART_OUTPUT_DIR), filename)

        fig, ax = plt.subplots(figsize=(10, 5.5))
        try:
            if chart_type == "bar":
                sns.barplot(
                    data=grouped, x=x_axis, y=y_axis,
                    color=sns.color_palette("viridis", n_colors=1)[0], ax=ax
                )
                ax.tick_params(axis="x", rotation=45)
            elif chart_type == "line":
                if x_axis == "YearMonth":
                    grouped = grouped.sort_values(x_axis)
                sns.lineplot(
                    data=grouped, x=x_axis, y=y_axis,
                    marker="o", ax=ax
                )
                ax.tick_params(axis="x", rotation=45)
            elif chart_type == "pie":
                if len(grouped) > 10:
                    grouped = grouped.nlargest(10, y_axis)
                # Biểu đồ tròn không phù hợp khi có giá trị âm hoặc tổng bằng 0.
                if (grouped[y_axis] < 0).any() or grouped[y_axis].sum() <= 0:
                    plt.close(fig)
                    return {
                        "error": "Không thể vẽ biểu đồ tròn khi có giá trị âm hoặc tổng không dương. Hãy dùng biểu đồ cột."
                    }
                ax.pie(
                    grouped[y_axis],
                    labels=grouped[x_axis].astype(str),
                    autopct="%1.1f%%",
                    startangle=140,
                )
                ax.axis("equal")

            ax.set_title(chart_title, fontsize=13, fontweight="bold")
            if chart_type != "pie":
                ax.set_xlabel(x_axis)
                ax.set_ylabel(f"{aggregation.title()} of {y_axis}")
            fig.tight_layout()
            fig.savefig(save_path, dpi=150, bbox_inches="tight")
        finally:
            plt.close(fig)

        return {
            "status": "success",
            "chart_title": chart_title,
            "chart_type": chart_type,
            "x_axis": x_axis,
            "y_axis": y_axis,
            "aggregation": aggregation,
            "chart_filename": filename,
            "chart_path": save_path,
            "groups_plotted": int(len(grouped)),
        }
    except Exception as exc:
        return {"error": f"Không thể tạo biểu đồ: {exc}"}


# =============================================================================
# Tool 4: Phân tích sản phẩm bán chạy/bán chậm và lỗ
# =============================================================================

@tool
def analyze_products(
    ranking_by: str = "Sales",
    ascending: bool = False,
    category: Optional[str] = None,
    region: Optional[str] = None,
    top_n: int = 10,
) -> dict:
    """Xếp hạng sản phẩm theo tổng Sales, Profit hoặc Quantity; dùng ascending=True để tìm nhóm thấp nhất và trả thêm trạng thái lãi/lỗ."""
    try:
        if ranking_by not in {"Sales", "Profit", "Quantity"}:
            return {"error": "ranking_by phải là Sales, Profit hoặc Quantity."}
        if top_n < 1 or top_n > 100:
            return {"error": "top_n phải nằm trong khoảng 1 đến 100."}

        data = _filter_data(
            _get_data(),
            category=category,
            region=region,
        )
        product_column = "Product Name" if "Product Name" in data.columns else (
            "Product ID" if "Product ID" in data.columns else None
        )
        if product_column is None:
            return {
                "error": "Không tìm thấy cột Product Name hoặc Product ID để phân tích sản phẩm."
            }
        if data.empty:
            return {"total_products": 0, "products": []}

        for col in ("Sales", "Profit", "Quantity"):
            if col in data.columns:
                data[col] = pd.to_numeric(data[col], errors="coerce")

        aggregations = {
            col: "sum" for col in ("Sales", "Profit", "Quantity")
            if col in data.columns
        }
        products = (
            data.groupby(product_column, dropna=False)
            .agg(aggregations)
            .reset_index()
        )
        if ranking_by not in products.columns:
            return {"error": f"Không có dữ liệu để xếp hạng theo {ranking_by}."}

        products = products.sort_values(
            ranking_by, ascending=ascending, na_position="last"
        )
        selected = products.head(top_n).copy()
        if "Profit" in selected.columns:
            selected["profit_status"] = selected["Profit"].apply(
                lambda value: "loss" if pd.notna(value) and value < 0
                else ("profitable" if pd.notna(value) else "unknown")
            )

        return {
            "ranking_by": ranking_by,
            "order": "ascending" if ascending else "descending",
            "interpretation": (
                "Nhóm có giá trị thấp nhất theo tiêu chí đã chọn."
                if ascending else "Nhóm có giá trị cao nhất theo tiêu chí đã chọn."
            ),
            "total_products": int(len(products)),
            "products": _json_safe_records(selected),
            "note": (
                "Bán chậm được xác định tương đối theo tổng chỉ số trong dữ liệu hiện có; "
                "để đánh giá tốc độ bán chính xác cần xét thêm khoảng thời gian, tồn kho và thời gian sản phẩm được bày bán."
            ),
        }
    except Exception as exc:
        return {"error": f"Không thể phân tích sản phẩm: {exc}"}

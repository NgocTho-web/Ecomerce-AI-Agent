from typing import Optional, Literal
from pydantic import BaseModel, Field

# ==============================================================================
# 1. SCHEMA DÀNH CHO CÔNG CỤ LỌC VÀ TRUY VẤN DỮ LIỆU (Filter Data Tool)
# ==============================================================================
class FilterSalesInput(BaseModel):
    category: Optional[Literal["Furniture", "Office Supplies", "Technology"]] = Field(
        default=None, 
        description="Danh mục ngành hàng chính (Furniture, Office Supplies, Technology)."
    )
    sub_category: Optional[str] = Field(
        default=None, 
        description="Danh mục chi tiết của sản phẩm (ví dụ: Phones, Chairs, Storage, Tables, v.v.)."
    )
    segment: Optional[Literal["Consumer", "Corporate", "Home Office"]] = Field(
        default=None, 
        description="Phân khúc khách hàng mục tiêu."
    )
    region: Optional[Literal["Central", "East", "South", "West"]] = Field(
        default=None, 
        description="Khu vực địa lý giao hàng tại Mỹ."
    )
    state: Optional[str] = Field(
        default=None, 
        description="Tên bang cụ thể (ví dụ: California, New York, Texas)."
    )
    year: Optional[int] = Field(
        default=None, 
        description="Năm đặt hàng (ví dụ: 2011, 2012, 2013, 2014)."
    )
    is_loss: Optional[bool] = Field(
        default=None, 
        description="Nếu True, chỉ lọc các đơn hàng bị lỗ (Profit < 0). Nếu False, chỉ lọc đơn có lãi (Profit >= 0)."
    )
    limit: int = Field(
        default=5, 
        description="Số lượng dòng kết quả tối đa cần trích xuất mẫu."
    )


# ==============================================================================
# 2. SCHEMA DÀNH CHO CÔNG CỤ TÍNH TOÁN THỐNG KÊ (Calculate Metrics Tool)
# ==============================================================================
class CalculateMetricsInput(BaseModel):
    metric: Literal["Sales", "Profit", "Quantity", "Discount"] = Field(
        description="Chỉ số định lượng cần tính toán."
    )
    aggregation: Literal["sum", "mean", "median", "min", "max", "count"] = Field(
        default="sum", 
        description="Phép toán thống kê cần áp dụng (tổng, trung bình, trung vị, lớn nhất, nhỏ nhất, số lượng)."
    )
    group_by: Optional[Literal["Category", "Sub-Category", "Segment", "Region", "State", "Ship Mode", "YearMonth"]] = Field(
        default=None, 
        description="Thuộc tính để phân nhóm thống kê (nếu muốn tính theo từng nhóm)."
    )
    category: Optional[Literal["Furniture", "Office Supplies", "Technology"]] = Field(
        default=None, 
        description="Điều kiện lọc theo ngành hàng trước khi thống kê."
    )
    segment: Optional[Literal["Consumer", "Corporate", "Home Office"]] = Field(
        default=None, 
        description="Điều kiện lọc theo phân khúc khách hàng trước khi thống kê."
    )
    region: Optional[Literal["Central", "East", "South", "West"]] = Field(
        default=None, 
        description="Điều kiện lọc theo vùng địa lý trước khi thống kê."
    )
    year: Optional[int] = Field(
        default=None, 
        description="Điều kiện lọc theo năm cụ thể trước khi thống kê."
    )


# ==============================================================================
# 3. SCHEMA DÀNH CHO CÔNG CỤ VẼ BIỂU ĐỒ (Generate Chart Tool)
# ==============================================================================
class GenerateChartInput(BaseModel):
    chart_type: Literal["bar", "line", "pie"] = Field(
        description="Loại biểu đồ trực quan hóa (cột, đường, tròn)."
    )
    x_axis: Literal["Category", "Sub-Category", "Segment", "Region", "YearMonth", "Ship Mode"] = Field(
        description="Trục phân loại hoặc trục thời gian (trục X)."
    )
    y_axis: Literal["Sales", "Profit", "Quantity"] = Field(
        default="Sales", 
        description="Chỉ số định lượng hiển thị trên trục Y."
    )
    aggregation: Literal["sum", "mean"] = Field(
        default="sum", 
        description="Phép tổng hợp cho giá trị hiển thị trên biểu đồ."
    )
    chart_title: str = Field(
        default="Biểu đồ phân tích kinh doanh", 
        description="Tiêu đề hiển thị của biểu đồ."
    )


# ==============================================================================
# 4. SCHEMA ĐẦU RA CHUẨN CỦA AGENT TRẢ VỀ CHO GIAO DIỆN STREAMLIT 
# ==============================================================================
class AgentResponseOutput(BaseModel):
    answer: str = Field(
        description="Câu trả lời tóm tắt ngắn gọn và giải thích số liệu bằng tiếng Việt cho người dùng."
    )
    chart_path: Optional[str] = Field(
        default=None, 
        description="Đường dẫn tương đối tới file ảnh biểu đồ được lưu (nếu yêu cầu có vẽ biểu đồ)."
    )
    raw_table: Optional[list[dict]] = Field(
        default=None, 
        description="Dữ liệu dạng danh sách bản ghi nếu người dùng yêu cầu xem chi tiết bảng kết quả."
    )
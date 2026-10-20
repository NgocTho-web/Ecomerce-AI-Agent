import os
from dotenv import load_dotenv

# Tự động nạp các biến môi trường từ file .env
load_dotenv()

# Cấu hình API Key (Hỗ trợ OpenAI hoặc Google Gemini)
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "")

# Đường dẫn đến tệp dữ liệu đã được Thành viên 2 tiền xử lý
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLEANED_DATA_PATH = os.path.join(BASE_DIR, "data", "processed", "ecommerce_cleaned_for_agent.csv")
CHART_OUTPUT_DIR = os.path.join(BASE_DIR, "app", "assets")

# Đảm bảo thư mục lưu biểu đồ tồn tại để Thành viên 4 load lên UI
os.makedirs(CHART_OUTPUT_DIR, exist_ok=True)
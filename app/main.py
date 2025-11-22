from fastapi import FastAPI
from motor.motor_asyncio import AsyncIOMotorClient
import os
import logging
from dotenv import load_dotenv

from app.routes import cong_van


# Load environment variables from a .env file
load_dotenv()


app = FastAPI()

# Cấu hình MongoDB
MONGODB_URL = os.getenv("MONGODB_URL", "mongodb://localhost:27017")
DATABASE_NAME = os.getenv("DATABASE_NAME", "d68cv")

# Khởi tạo MongoDB client
client = AsyncIOMotorClient(MONGODB_URL)
db = client[DATABASE_NAME]

# Expose the db on the FastAPI app state so routes can access it via Request
app.state.db = db


# Cấu hình logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


@app.on_event("startup")
async def startup_db_client():
    """Kết nối database khi khởi động"""
    try:
        await client.admin.command("ping")
        print("✅ Kết nối MongoDB thành công!")
    except Exception as e:
        print(f"❌ Lỗi kết nối MongoDB: {e}")


@app.on_event("shutdown")
async def shutdown_db_client():
    """Đóng kết nối database khi tắt"""
    client.close()


# default route
@app.get("/")
def read_root():
    return {"status": "API is running"}


# routes
app.include_router(cong_van.router, prefix="/api/cong-van", tags=["cong-van"])


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=os.getenv("DEFAULT_PORT", "8080"), reload=True)

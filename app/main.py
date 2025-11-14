from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from motor.motor_asyncio import AsyncIOMotorClient
import pandas as pd
from io import BytesIO
from typing import List
import os
from dotenv import load_dotenv
from pydantic import BaseModel


# Load environment variables from a .env file
load_dotenv()


app = FastAPI()

# Cấu hình MongoDB
MONGODB_URL = os.getenv("MONGODB_URL", "mongodb://localhost:27017")
DATABASE_NAME = os.getenv("DATABASE_NAME", "d68cv_db")

# Khởi tạo MongoDB client
client = AsyncIOMotorClient(MONGODB_URL)
db = client[DATABASE_NAME]
cv_collection = db["cong_van"]


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


class CongVan(BaseModel):
    so_cong_van: str
    ngay_cong_van: str
    don_vi_giao: str
    sdt_lien_he: str
    dia_chi: str
    email: str


class ResponseModel(BaseModel):
    status: str
    data: List[CongVan]
    message: str


@app.get("/")
def read_root():
    return {"Hello": "World"}


@app.get("/cong-van/", tags=["cong-van"])
async def get_cong_van() -> dict:
    try:
        cong_van_stmt = cv_collection.find()
        cong_van_list = await cong_van_stmt.to_list(length=100)
        return ResponseModel(
            status="success",
            data=cong_van_list,
            message="Cong van retrieved successfully",
        ).dict()
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error retrieving cong van: {str(e)}"
        )


@app.post("/cong-van/", tags=["cong-van"])
async def create_cong_van(cong_van: CongVan) -> dict:
    try:
        cong_van_dict = cong_van.dict()
        result = await cv_collection.insert_one(cong_van_dict)
        cong_van_dict["_id"] = str(result.inserted_id)
        return ResponseModel(
            status="success",
            data=[cong_van_dict],
            message="Cong van created successfully",
        ).dict()
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error creating cong van: {str(e)}"
        )


@app.post("/upload-excel/")
async def upload_excel(
    file: UploadFile = File(...),
    collection_name: str = "cong_van",
    sheet_name: str = None,
):
    """
    Upload file Excel và import vào MongoDB

    Parameters:
    - file: File Excel (.xlsx, .xls)
    - collection_name: Tên collection trong MongoDB
    - sheet_name: Tên sheet cần import (None = sheet đầu tiên)
    """

    # Kiểm tra định dạng file
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400, detail="Chỉ chấp nhận file Excel (.xlsx, .xls)"
        )

    try:
        # Đọc file Excel
        contents = await file.read()
        excel_file = BytesIO(contents)

        # Đọc Excel với pandas
        if sheet_name:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
        else:
            df = pd.read_excel(excel_file)

        # Chuyển đổi NaN thành None
        df = df.where(pd.notna(df), None)

        # Chuyển DataFrame thành list of dictionaries
        records = df.to_dict("records")

        if not records:
            raise HTTPException(status_code=400, detail="File Excel không có dữ liệu")

        # Insert vào MongoDB
        collection = db[collection_name]
        result = await collection.insert_many(records)

        return JSONResponse(
            status_code=200,
            content={
                "message": "Import thành công",
                "collection": collection_name,
                "total_records": len(result.inserted_ids),
                "inserted_ids": [
                    str(id) for id in result.inserted_ids[:10]
                ],  # 10 ID đầu tiên
                "columns": list(df.columns),
                "sample_data": records[:3],  # 3 bản ghi mẫu
            },
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi xử lý file: {str(e)}")


# app.include_router(cong_van.router, prefix="/api/cong-van", tags=["cong-van"])


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)

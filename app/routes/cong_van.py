from fastapi import APIRouter, File, UploadFile, HTTPException, Depends
from fastapi.responses import JSONResponse
import pandas as pd
from typing import List, Optional, Any
from app.dependencies import get_db
from dotenv import load_dotenv
from pydantic import BaseModel
from unidecode import unidecode
import io
import re
import datetime


router = APIRouter()


# models
class CongVanModel(BaseModel):
    so_den: int
    so_cong_van: str
    ngay_cong_van: datetime.datetime
    don_vi_giao: Optional[str] = None
    sdt_lien_he: Optional[str] = None
    dia_chi: Optional[str] = None
    email: Optional[str] = None
    loi: Optional[str] = None
    ghi_chu: Optional[str] = None


class CongVanResponseModel(BaseModel):
    status: str
    data: List[CongVanModel]
    message: str


# routes
@router.get("/", tags=["cong-van"])
async def get_cong_van(db: Any = Depends(get_db)) -> dict:
    cv_collection = db["cong_van"]
    try:
        cong_van_stmt = cv_collection.find()
        cong_van_list = await cong_van_stmt.to_list(length=100)
        return CongVanResponseModel(
            status="success",
            data=cong_van_list,
            message="Cong van retrieved successfully",
        ).dict()
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error retrieving cong van: {str(e)}"
        )


@router.post("/", tags=["cong-van"])
async def create_cong_van(cong_van: CongVanModel, db: Any = Depends(get_db)) -> dict:
    cv_collection = db["cong_van"]
    try:
        cong_van_dict = cong_van.dict()
        result = await cv_collection.insert_one(cong_van_dict)
        cong_van_dict["_id"] = str(result.inserted_id)
        return CongVanResponseModel(
            status="success",
            data=[cong_van_dict],
            message="Cong van created successfully",
        ).dict()
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error creating cong van: {str(e)}"
        )


# Hàm chuyển tên cột sang snake_case
def to_snake_case(text):
    text = unidecode(text)  # bỏ dấu tiếng Việt
    text = text.strip().lower()
    text = re.sub(r"[^\w\s]", "", text)  # xóa ký tự đặc biệt
    text = re.sub(r"\s+", "_", text)  # đổi spaces → underscore
    return text


@router.post("/import-excel/", tags=["cong-van"])
async def import_excel(
    db: Any = Depends(get_db),
    file: UploadFile = File(...),
):
    cv_collection = db["cong_van"]

    try:
        # Đọc file Excel
        data_frame = pd.read_excel(file.file)
        data_frame = data_frame.fillna("")  # Thay thế giá trị NaN bằng chuỗi rỗng

        # Chuyển tất cả cột sang snake_case
        data_frame.columns = [to_snake_case(col) for col in data_frame.columns]

        # Chuyển DataFrame thành list dict
        records = data_frame.to_dict(orient="records")

        # Insert vào MongoDB
        if records:
            cv_collection.insert_many(records)

            return JSONResponse(
                status_code=200,
                content={
                    "status": "success",
                    "inserted_count": len(records),
                    "message": "Excel data imported successfully",
                },
            )
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error importing Excel data: {str(e)}"
        )


@router.delete("/{cong_van_id}", tags=["cong-van"])
async def delete_cong_van(
    cong_van_id: str, db: Any = Depends(get_db)
) -> dict:
    cv_collection = db["cong_van"]
    try:
        result = await cv_collection.delete_one({"_id": cong_van_id})
        if result.deleted_count == 1:
            return JSONResponse(
                status_code=200,
                content={
                    "status": "success",
                    "message": f"Cong van with id {cong_van_id} deleted successfully",
                },
            )
        else:
            raise HTTPException(
                status_code=404, detail=f"Cong van with id {cong_van_id} not found"
            )
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error deleting cong van: {str(e)}"
        )
    

@router.delete("/", tags=["cong-van"])
async def delete_all_cong_van(db: Any = Depends(get_db)) -> dict:
    cv_collection = db["cong_van"]
    try:
        result = await cv_collection.delete_many({})
        return JSONResponse(
            status_code=200,
            content={
                "status": "success",
                "deleted_count": result.deleted_count,
                "message": "All cong van records deleted successfully",
            },
        )
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error deleting cong van records: {str(e)}"
        )
    
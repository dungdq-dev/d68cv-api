from pydantic import BaseModel

class CongVan(BaseModel):
    so_cong_van: str
    ngay_cong_van: str
    don_vi_giao: str
    sdt_lien_he: str
    dia_chi: str
    email: str
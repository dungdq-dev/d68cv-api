from fastapi import APIRouter, FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
import pandas as pd
from io import BytesIO
from typing import List
import os
from dotenv import load_dotenv

app = FastAPI()
router = APIRouter()


app.include_router(router)

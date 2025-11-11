from fastapi import APIRouter, Depends
from app.schemas.user import User, UserCreate
from app.services.user_service import UserService

router = APIRouter()

@router.get("/users/{user_id}", response_model=User)
async def get_user(user_id: int):
    return await UserService.get_user(user_id)
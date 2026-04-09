from typing import List
from fastapi import APIRouter, Depends, Query
from sqlmodel import select
from sqlalchemy.ext.asyncio import AsyncSession
from app.database import get_session
from app.models.recipe import Recipe, RecipeRead, RecipeCategory

router = APIRouter()

@router.get("/", response_model=List[RecipeRead])
async def read_recipes(
    offset: int = 0,
    limit: int = Query(default=100, lte=100),
    session: AsyncSession = Depends(get_session)
):
    result = await session.execute(select(Recipe).offset(offset).limit(limit))
    return result.scalars().all()

@router.get("/categories")
async def read_categories(session: AsyncSession = Depends(get_session)):
    result = await session.execute(select(RecipeCategory))
    return result.scalars().all()

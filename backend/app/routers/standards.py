from typing import List
from fastapi import APIRouter, Depends, Query
from sqlmodel import select
from sqlalchemy.ext.asyncio import AsyncSession
from app.database import get_session
from app.models.layer_standard import LayerStandard, LayerStandardRead
from app.models.filename_rule import FilenameRule, FilenameRuleRead

router = APIRouter()

@router.get("/layers", response_model=List[LayerStandardRead])
async def read_layer_standards(
    offset: int = 0,
    limit: int = Query(default=100, lte=1000),
    category: str = None,
    session: AsyncSession = Depends(get_session)
):
    statement = select(LayerStandard)
    if category:
        statement = statement.where(LayerStandard.category == category)
    result = await session.execute(statement.offset(offset).limit(limit))
    return result.scalars().all()

@router.get("/filename-rules", response_model=List[FilenameRuleRead])
async def read_filename_rules(
    session: AsyncSession = Depends(get_session)
):
    result = await session.execute(select(FilenameRule))
    return result.scalars().all()

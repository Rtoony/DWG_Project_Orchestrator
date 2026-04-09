from typing import List, Optional
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlmodel import select
from sqlalchemy.ext.asyncio import AsyncSession
from app.database import get_session
from app.models.job import Job, JobRead

router = APIRouter()

@router.post("/", response_model=JobRead)
async def create_job(job: Job, session: AsyncSession = Depends(get_session)):
    session.add(job)
    await session.commit()
    await session.refresh(job)
    return job

@router.get("/", response_model=List[JobRead])
async def read_jobs(
    project_id: Optional[int] = None,
    status: Optional[str] = None,
    offset: int = 0,
    limit: int = Query(default=100, lte=100),
    session: AsyncSession = Depends(get_session)
):
    statement = select(Job)
    if project_id:
        statement = statement.where(Job.project_id == project_id)
    if status:
        statement = statement.where(Job.status == status)
    result = await session.execute(statement.offset(offset).limit(limit))
    return result.scalars().all()

@router.get("/{job_id}", response_model=JobRead)
async def read_job(job_id: int, session: AsyncSession = Depends(get_session)):
    job = await session.get(Job, job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    return job

from pydantic import BaseModel
from typing import Any, Dict
from datetime import datetime

class JobUpdate(BaseModel):
    status: Optional[str] = None
    progress: Optional[int] = None
    result: Optional[Dict[str, Any]] = None
    worker_id: Optional[str] = None
    started_at: Optional[datetime] = None
    completed_at: Optional[datetime] = None

@router.put("/{job_id}", response_model=JobRead)
async def update_job(job_id: int, job_update: JobUpdate, session: AsyncSession = Depends(get_session)):
    job = await session.get(Job, job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    
    update_data = job_update.model_dump(exclude_unset=True)
    for key, value in update_data.items():
        setattr(job, key, value)
        
    session.add(job)
    await session.commit()
    await session.refresh(job)
    return job

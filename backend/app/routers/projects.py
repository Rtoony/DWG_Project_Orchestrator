from typing import List, Optional
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlmodel import select
from sqlalchemy.ext.asyncio import AsyncSession
from app.database import get_session
from app.models.project import Project, ProjectCreate, ProjectRead

router = APIRouter()

@router.post("/", response_model=ProjectRead)
async def create_project(project: ProjectCreate, session: AsyncSession = Depends(get_session)):
    db_project = Project.from_orm(project)
    session.add(db_project)
    await session.commit()
    await session.refresh(db_project)
    return db_project

@router.get("/", response_model=List[ProjectRead])
async def read_projects(
    offset: int = 0,
    limit: int = Query(default=100, lte=100),
    session: AsyncSession = Depends(get_session)
):
    result = await session.execute(select(Project).offset(offset).limit(limit))
    projects = result.scalars().all()
    return projects

@router.get("/{project_id}", response_model=ProjectRead)
async def read_project(project_id: int, session: AsyncSession = Depends(get_session)):
    project = await session.get(Project, project_id)
    if not project:
        raise HTTPException(status_code=404, detail="Project not found")
    return project

@router.patch("/{project_id}", response_model=ProjectRead)
async def update_project(
    project_id: int, project: ProjectCreate, session: AsyncSession = Depends(get_session)
):
    db_project = await session.get(Project, project_id)
    if not db_project:
        raise HTTPException(status_code=404, detail="Project not found")
    project_data = project.dict(exclude_unset=True)
    for key, value in project_data.items():
        setattr(db_project, key, value)
    session.add(db_project)
    await session.commit()
    await session.refresh(db_project)
    return db_project

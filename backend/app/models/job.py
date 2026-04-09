from datetime import datetime
from typing import Optional, Dict, Any
from sqlmodel import SQLModel, Field, Column, JSON, Relationship

class JobBase(SQLModel):
    project_id: Optional[int] = Field(default=None, foreign_key="projects.id")
    job_type: str
    status: str = Field(default="pending")
    payload: Dict[str, Any] = Field(default={}, sa_column=Column(JSON))
    progress: int = Field(default=0)
    worker_id: Optional[str] = None

class Job(JobBase, table=True):
    __tablename__ = "jobs"
    id: Optional[int] = Field(default=None, primary_key=True)
    result: Optional[Dict[str, Any]] = Field(default=None, sa_column=Column(JSON))
    created_at: datetime = Field(default_factory=datetime.utcnow)
    started_at: Optional[datetime] = None
    completed_at: Optional[datetime] = None

    project: Optional["Project"] = Relationship(back_populates="jobs")

class JobRead(JobBase):
    id: int
    result: Optional[Dict[str, Any]]
    created_at: datetime
    started_at: Optional[datetime]
    completed_at: Optional[datetime]

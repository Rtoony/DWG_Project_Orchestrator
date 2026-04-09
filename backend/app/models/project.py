from datetime import datetime
from typing import Optional, List
from sqlmodel import SQLModel, Field, Relationship

class ProjectBase(SQLModel):
    project_number: str = Field(index=True)
    sub_number: str = Field(index=True)
    project_name: Optional[str] = None
    client_name: Optional[str] = None
    project_manager: Optional[str] = None
    lead_designer: Optional[str] = None
    project_date: Optional[datetime] = None
    project_status: str = Field(default="SD")
    setup_config: Optional[str] = None
    tb_size: Optional[str] = None
    tb_type: Optional[str] = None
    coordinate_system: Optional[str] = None
    vertical_datum: Optional[str] = None
    root_path: str = Field(default="J:\J")
    archive_path: str = Field(default="R:\J")

class Project(ProjectBase, table=True):
    __tablename__ = "projects"
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)
    
    drawings: List["Drawing"] = Relationship(back_populates="project")
    jobs: List["Job"] = Relationship(back_populates="project")
    audit_logs: List["AuditLog"] = Relationship(back_populates="project")

class ProjectCreate(ProjectBase):
    pass

class ProjectRead(ProjectBase):
    id: int
    created_at: datetime
    updated_at: datetime

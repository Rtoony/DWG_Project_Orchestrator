from datetime import datetime
from typing import Optional, List
from sqlmodel import SQLModel, Field, Relationship

class DrawingBase(SQLModel):
    project_id: int = Field(foreign_key="projects.id")
    filename: str
    file_type_code: str
    description: Optional[str] = None
    phase: Optional[str] = None
    folder_path: Optional[str] = None
    file_size_bytes: Optional[int] = None
    dwg_version: Optional[str] = None
    last_modified: Optional[datetime] = None
    status: str = Field(default="active")

class Drawing(DrawingBase, table=True):
    __tablename__ = "drawings"
    id: Optional[int] = Field(default=None, primary_key=True)
    last_analyzed: Optional[datetime] = None
    created_at: datetime = Field(default_factory=datetime.utcnow)

    project: "Project" = Relationship(back_populates="drawings")
    analyses: List["DXFAnalysis"] = Relationship(back_populates="drawing")

class DrawingCreate(DrawingBase):
    pass

class DrawingRead(DrawingBase):
    id: int
    last_analyzed: Optional[datetime]
    created_at: datetime

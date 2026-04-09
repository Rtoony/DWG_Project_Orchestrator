from typing import List, Dict, Any, Optional
from sqlmodel import SQLModel, Field, Column, JSON

class ProjectPresetBase(SQLModel):
    name: str = Field(unique=True, index=True)
    description: Optional[str] = None
    drawings: List[Dict[str, Any]] = Field(default=[], sa_column=Column(JSON))

class ProjectPreset(ProjectPresetBase, table=True):
    __tablename__ = "project_presets"
    id: Optional[int] = Field(default=None, primary_key=True)

class ProjectPresetCreate(ProjectPresetBase):
    pass

class ProjectPresetRead(ProjectPresetBase):
    id: int

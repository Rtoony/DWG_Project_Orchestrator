from datetime import datetime
from typing import Optional, Dict, Any
from sqlmodel import SQLModel, Field, Column, JSON, Relationship

class DXFAnalysisBase(SQLModel):
    drawing_id: int = Field(foreign_key="drawings.id")
    analysis_data: Dict[str, Any] = Field(default={}, sa_column=Column(JSON))
    entity_count: Optional[int] = None
    layer_count: Optional[int] = None
    block_count: Optional[int] = None

class DXFAnalysis(DXFAnalysisBase, table=True):
    __tablename__ = "dxf_analyses"
    id: Optional[int] = Field(default=None, primary_key=True)
    analyzed_at: datetime = Field(default_factory=datetime.utcnow)
    
    drawing: "Drawing" = Relationship(back_populates="analyses")

class DXFAnalysisRead(DXFAnalysisBase):
    id: int
    analyzed_at: datetime

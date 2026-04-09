from datetime import datetime
from typing import Optional, List
from sqlmodel import SQLModel, Field, Column, JSON

class LayerStandardBase(SQLModel):
    name: str = Field(unique=True, index=True)
    color_code: int
    linetype: str = Field(default="CONTINUOUS")
    lineweight: int = Field(default=-3)
    is_plottable: bool = Field(default=True)
    plot_style_name: Optional[str] = None
    category: Optional[str] = None
    discipline: Optional[str] = None
    status: Optional[str] = None
    description: Optional[str] = None
    typical_object_types: List[str] = Field(default=[], sa_column=Column(JSON))
    notes: Optional[str] = None
    standards_revision_id: int = Field(default=1)

class LayerStandard(LayerStandardBase, table=True):
    __tablename__ = "layer_standards"
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)

class LayerStandardCreate(LayerStandardBase):
    pass

class LayerStandardRead(LayerStandardBase):
    id: int
    created_at: datetime
    updated_at: datetime

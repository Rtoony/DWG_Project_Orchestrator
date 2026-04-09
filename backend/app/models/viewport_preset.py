from typing import List, Dict, Any, Optional
from sqlmodel import SQLModel, Field, Column, JSON

class ViewportPresetBase(SQLModel):
    tb_type: str
    tb_size: str
    layout_code: str
    viewports: List[Dict[str, Any]] = Field(default=[], sa_column=Column(JSON))

class ViewportPreset(ViewportPresetBase, table=True):
    __tablename__ = "viewport_presets"
    id: Optional[int] = Field(default=None, primary_key=True)

class ViewportPresetCreate(ViewportPresetBase):
    pass

class ViewportPresetRead(ViewportPresetBase):
    id: int

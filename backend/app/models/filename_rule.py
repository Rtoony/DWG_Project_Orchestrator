from datetime import datetime
from typing import Optional
from sqlmodel import SQLModel, Field

class FilenameRuleBase(SQLModel):
    file_type_code: str = Field(unique=True, index=True)
    file_type_description: Optional[str] = None
    folder_path_template: Optional[str] = None
    filename_pattern: Optional[str] = None
    phase_required: bool = Field(default=False)
    phase_source: Optional[str] = None
    phase_allowed_list_source: Optional[str] = None
    phase_format: Optional[str] = None
    description_required: bool = Field(default=False)
    description_format: Optional[str] = None
    multi_instance_allowed: bool = Field(default=True)
    notes: Optional[str] = None

class FilenameRule(FilenameRuleBase, table=True):
    __tablename__ = "filename_rules"
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)

class FilenameRuleCreate(FilenameRuleBase):
    pass

class FilenameRuleRead(FilenameRuleBase):
    id: int
    created_at: datetime

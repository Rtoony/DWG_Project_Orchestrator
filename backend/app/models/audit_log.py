from datetime import datetime
from typing import Optional, Dict, Any
from sqlmodel import SQLModel, Field, Column, JSON, Relationship

class AuditLogBase(SQLModel):
    project_id: Optional[int] = Field(default=None, foreign_key="projects.id")
    action: str
    details: Dict[str, Any] = Field(default={}, sa_column=Column(JSON))
    user_name: Optional[str] = None

class AuditLog(AuditLogBase, table=True):
    __tablename__ = "audit_log"
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)

    project: Optional["Project"] = Relationship(back_populates="audit_logs")

class AuditLogRead(AuditLogBase):
    id: int
    created_at: datetime

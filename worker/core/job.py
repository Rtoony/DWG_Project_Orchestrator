from enum import Enum
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional
from pydantic import BaseModel

class JobKind(str, Enum):
    ACCORECONSOLE = "accoreconsole"
    COM = "com"
    PYTHON = "python"

@dataclass
class FileResult:
    dwg_path: str
    status: str  # "success", "error", "skipped"
    duration_ms: int = 0
    stdout: str = ""
    stderr: str = ""
    error: Optional[str] = None
    sentinel_found: bool = False

class BaseJob:
    name: str = "base-job"
    kind: JobKind = JobKind.ACCORECONSOLE
    description: str = ""
    template: Optional[str] = None  # For accoreconsole scripts

    class Inputs(BaseModel):
        pass

    def parse_result(self, stdout: str, stderr: str) -> bool:
        """Return True if the job succeeded based on output."""
        return "NEXUS_OK" in stdout or "NEXUS_OK" in stderr

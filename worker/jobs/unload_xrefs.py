from core.job import BaseJob, JobKind
from pydantic import BaseModel
from typing import Literal

class UnloadXrefsJob(BaseJob):
    name = "unload-xrefs"
    kind = JobKind.ACCORECONSOLE
    description = "Bulk unload, detach, or reload all external references"
    template = "unload_xrefs.scr.j2"

    class Inputs(BaseModel):
        mode: Literal["Unload", "Detach", "Reload"] = "Unload"
        save_in_place: bool = True

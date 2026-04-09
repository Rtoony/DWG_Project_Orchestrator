from core.job import BaseJob, JobKind
from pydantic import BaseModel

class PurgeJob(BaseJob):
    name = "purge-all"
    kind = JobKind.ACCORECONSOLE
    description = "Structural purge (blocks, layers, regapps) + AUDIT recovery"
    template = "purge_all.scr.j2"

    class Inputs(BaseModel):
        run_audit: bool = True
        purge_regapps: bool = True
        save_in_place: bool = True

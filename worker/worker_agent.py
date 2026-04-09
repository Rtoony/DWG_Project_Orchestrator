import asyncio
import time
import requests
import yaml
import logging
from datetime import datetime
import sys
from pathlib import Path
import importlib
import pkgutil
import uuid

from accoreconsole_runner import ParallelRunner
from autocad_engine import AutoCADEngine
from core.job import BaseJob, JobKind

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("WorkerAgent")

class WorkerAgent:
    def __init__(self, config_path: str = "config.yaml"):
        with open(config_path, "r") as f:
            self.config = yaml.safe_load(f)
            
        self.api_url = self.config.get("api_url", "http://localhost:8000/api/v1")
        self.worker_id = self.config.get("worker_id", "cad_worker_01")
        self.poll_interval = self.config.get("poll_interval_seconds", 5)
        self.max_concurrent = self.config.get("max_concurrent", 4)
        
        self.jobs = self._discover_jobs()
        self.parallel_runner = ParallelRunner(max_concurrent=self.max_concurrent)
        self.acad_engine = None
        
        logger.info(f"Worker initialized: ID={self.worker_id}, API={self.api_url}, Jobs={list(self.jobs.keys())}")

    def _discover_jobs(self) -> dict:
        jobs = {}
        jobs_path = Path(__file__).parent / "jobs"
        for _, name, _ in pkgutil.iter_modules([str(jobs_path)]):
            module = importlib.import_module(f"jobs.{name}")
            for attr in dir(module):
                cls = getattr(module, attr)
                if isinstance(cls, type) and issubclass(cls, BaseJob) and cls != BaseJob:
                    jobs[cls.name] = cls()
        return jobs

    async def start(self):
        logger.info("Starting CAD Worker Agent...")
        while True:
            try:
                job_data = self.fetch_job()
                if job_data:
                    await self.process_job(job_data)
                else:
                    await asyncio.sleep(self.poll_interval)
            except KeyboardInterrupt:
                logger.info("Worker stopped by user")
                break
            except Exception as e:
                logger.error(f"Unexpected error in main loop: {e}")
                await asyncio.sleep(self.poll_interval)

    def fetch_job(self) -> dict:
        try:
            response = requests.get(f"{self.api_url}/jobs?status=pending&limit=1")
            if response.status_code == 200:
                jobs = response.json()
                if jobs and len(jobs) > 0:
                    job = jobs[0]
                    update_data = {
                        "status": "running",
                        "worker_id": self.worker_id,
                        "started_at": datetime.utcnow().isoformat()
                    }
                    claim_resp = requests.put(f"{self.api_url}/jobs/{job['id']}", json=update_data)
                    if claim_resp.status_code == 200:
                        return claim_resp.json()
        except requests.exceptions.RequestException:
            pass
        return None

    def update_job(self, job_id: int, status: str, result: dict = None, progress: int = None):
        try:
            update_data = {"status": status}
            if result is not None: update_data["result"] = result
            if progress is not None: update_data["progress"] = progress
            if status in ["completed", "failed"]:
                update_data["completed_at"] = datetime.utcnow().isoformat()
            requests.put(f"{self.api_url}/jobs/{job_id}", json=update_data)
        except Exception as e:
            logger.error(f"Failed to update job {job_id}: {e}")

    async def process_job(self, job_data: dict):
        job_id = job_data.get("id")
        payload = job_data.get("payload", {})
        job_name = payload.get("job_name")
        dwg_paths = payload.get("dwg_paths", [])
        job_inputs = payload.get("inputs", {})
        run_id = str(uuid.uuid4())[:8]

        if job_name not in self.jobs:
            self.update_job(job_id, "failed", result={"error": f"Job {job_name} not found in worker registry"})
            return

        job_obj = self.jobs[job_name]
        logger.info(f"Processing job {job_id}: {job_name} on {len(dwg_paths)} files")

        try:
            if job_obj.kind == JobKind.ACCORECONSOLE:
                results = await self.parallel_runner.run_batch(
                    job_name, job_obj.template, dwg_paths, job_inputs, run_id
                )
                # Check for sentinel/success in each result
                for r in results:
                    r["success"] = job_obj.parse_result(r.get("stdout", ""), r.get("stderr", ""))
                
                failed_count = sum(1 for r in results if not r.get("success"))
                status = "completed" if failed_count == 0 else "failed"
                self.update_job(job_id, status, result={"files": results}, progress=100)
            
            elif job_obj.kind == JobKind.COM:
                # COM jobs are currently serial to prevent AutoCAD GUI collision
                if not self.acad_engine: self.acad_engine = AutoCADEngine()
                results = []
                for dwg in dwg_paths:
                    # In a real impl, we'd call job_obj.run_com(doc, inputs)
                    # For now, we'll use our existing run_command
                    res = self.acad_engine.run_command(dwg, job_inputs.get("command", "QSAVE"))
                    results.append(res)
                self.update_job(job_id, "completed", result={"files": results}, progress=100)

        except Exception as e:
            logger.error(f"Error processing job {job_id}: {e}")
            self.update_job(job_id, "failed", result={"error": str(e)})

if __name__ == "__main__":
    agent = WorkerAgent()
    asyncio.run(agent.start())

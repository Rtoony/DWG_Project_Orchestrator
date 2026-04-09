import os
import subprocess
import logging
import asyncio
import time
from pathlib import Path
import tempfile
from jinja2 import Environment, FileSystemLoader, StrictUndefined

logger = logging.getLogger(__name__)

# Config for Jinja2
TEMPLATE_DIR = Path(__file__).parent / "templates"
jinja_env = Environment(
    loader=FileSystemLoader(str(TEMPLATE_DIR)),
    undefined=StrictUndefined,
    keep_trailing_newline=True
)

def find_accoreconsole(prefer_year="2026") -> Path:
    env_path = os.environ.get("ACCORECONSOLE_EXE")
    if env_path and Path(env_path).exists():
        return Path(env_path)
    
    roots = [Path(r"C:\Program Files\Autodesk"), Path(r"C:\Program Files (x86)\Autodesk")]
    for root in roots:
        if root.exists():
            for p in root.rglob("accoreconsole.exe"):
                if prefer_year in str(p):
                    return p
    # Fallback to a common path if not found
    return Path(fr"C:\Program Files\Autodesk\AutoCAD {prefer_year}\accoreconsole.exe")

async def run_accore_single(dwg_path: str, script_path: str, timeout: int = 300) -> dict:
    exe_path = find_accoreconsole()
    
    args = [
        str(exe_path),
        "/i", str(dwg_path),
        "/s", str(script_path),
        "/l", "en-US"
    ]
    
    start_time = time.time()
    try:
        proc = await asyncio.create_subprocess_exec(
            *args,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE
        )
        
        try:
            stdout, stderr = await asyncio.wait_for(proc.communicate(), timeout=timeout)
            duration = int((time.time() - start_time) * 1000)
            
            stdout_str = stdout.decode('utf-8', errors='ignore')
            stderr_str = stderr.decode('utf-8', errors='ignore')
            
            return {
                "status": "success" if proc.returncode == 0 else "error",
                "returncode": proc.returncode,
                "stdout": stdout_str,
                "stderr": stderr_str,
                "duration_ms": duration
            }
        except asyncio.TimeoutError:
            proc.kill()
            return {"status": "error", "error": "Timeout expired", "duration_ms": int((time.time() - start_time) * 1000)}
            
    except Exception as e:
        return {"status": "error", "error": str(e)}

class ParallelRunner:
    def __init__(self, max_concurrent: int = 4):
        self.semaphore = asyncio.Semaphore(max_concurrent)

    async def run_batch(self, job_name: str, template_name: str, dwg_paths: list[str], inputs: dict, run_id: str):
        tasks = []
        for dwg in dwg_paths:
            tasks.append(self._run_with_semaphore(job_name, template_name, dwg, inputs, run_id))
        return await asyncio.gather(*tasks)

    async def _run_with_semaphore(self, job_name: str, template_name: str, dwg: str, inputs: dict, run_id: str):
        async with self.semaphore:
            # Render template
            template = jinja_env.get_template(template_name)
            render_ctx = {**inputs, "run_id": run_id, "job_name": job_name}
            script_content = template.render(**render_ctx)
            
            with tempfile.NamedTemporaryFile(mode='w', suffix='.scr', delete=False) as f:
                f.write(script_content)
                temp_script = f.name
                
            try:
                logger.info(f"Starting {job_name} on {Path(dwg).name}")
                result = await run_accore_single(dwg, temp_script)
                result["dwg_path"] = dwg
                return result
            finally:
                try: os.remove(temp_script)
                except: pass

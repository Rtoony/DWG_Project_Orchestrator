import time
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

class AutoCADEngine:
    def __init__(self):
        self.app = None
        self._connected = False

    def connect(self):
        try:
            import pythoncom
            from win32com.client.gencache import EnsureDispatch
            pythoncom.CoInitialize()
            self.app = EnsureDispatch("AutoCAD.Application")
            self.app.Visible = True
            self._connected = True
            logger.info("Successfully connected to AutoCAD")
            return True
        except Exception as e:
            logger.error(f"Failed to connect to AutoCAD: {e}")
            self._connected = False
            return False

    def wait_quiet(self, timeout: float = 60.0):
        if not self.app: return
        t0 = time.time()
        st = self.app.GetAcadState()
        while not st.IsQuiescent:
            if time.time() - t0 > timeout:
                raise TimeoutError("AutoCAD stayed busy for too long")
            time.sleep(0.25)
            st = self.app.GetAcadState()

    def run_command(self, dwg_path: str, command: str) -> dict:
        if not self._connected:
            if not self.connect():
                return {"status": "error", "error": "Could not connect to AutoCAD"}
        
        try:
            doc = None
            path = Path(dwg_path).resolve()
            
            # Find if open
            for d in self.app.Documents:
                if Path(d.FullName).resolve() == path:
                    doc = d
                    doc.Activate()
                    break
                    
            if not doc:
                doc = self.app.Documents.Open(str(path))
                doc.Activate()
                
            self.wait_quiet()
            
            cmd_str = command if command.endswith("\n") else f"{command}\n"
            doc.SendCommand(cmd_str)
            self.wait_quiet()
            
            return {"status": "success", "message": f"Executed {command} on {path.name}"}
        except Exception as e:
            return {"status": "error", "error": str(e)}

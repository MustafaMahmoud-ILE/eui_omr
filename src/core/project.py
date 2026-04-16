import json
import shutil
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List

from src.models.schemas import AnswerKey, GradingResult

class ProjectManager:
    """Handles persistence, state saving, and file integrity for an OMR project."""

    def __init__(self, project_dir: str | Path):
        self.project_dir = Path(project_dir)
        self.state_file = self.project_dir / "project_state.json"
        # Directory for cached OMR crops
        self.crops_dir = self.project_dir / "crops"
        
        # State variables
        self.project_name: str = "Untitled Project"
        self.question_count: int = 60
        self.active_models: int = 6 # Up to 6
        self.excel_roster_path: str = ""
        self.original_excel_path: str = ""
        self.student_id_col: str = ""
        self.grade_output_col: str = ""
        self.student_pdf_path: str = ""
        self.answer_keys: Dict[str, AnswerKey] = {}
        self.last_results: List[GradingResult] = []
        self.mark_sensitivity: int = 75 # 0-100 scale, default 75 (25% threshold)

        self.logger = logging.getLogger("ProjectManager")
        self._setup_logging()

        # Load state if it exists
        if self.state_file.exists():
            self.load_state()

    def _setup_logging(self):
        """Initializes a log file within the project directory."""
        log_path = self.project_dir / "project.log"
        self.logger.setLevel(logging.DEBUG)
        
        # Avoid duplicate handlers if re-initialized
        if not self.logger.handlers:
            fh = logging.FileHandler(log_path, encoding='utf-8')
            fh.setLevel(logging.DEBUG)
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            fh.setFormatter(formatter)
            self.logger.addHandler(fh)
        
        self.logger.info(f"--- Project Session Started: {self.project_name} ---")

    def create_project(self, name: str):
        self.project_name = name
        self.project_dir.mkdir(parents=True, exist_ok=True)
        # Create an 'imports' subfolder to safeguard local copies of user files
        (self.project_dir / "imports").mkdir(exist_ok=True)
        # Create 'crops' subfolder for caching vision results
        self.crops_dir.mkdir(exist_ok=True)
        self.save_state()

    def import_excel_file(self, original_file_path: str | Path) -> Path:
        """Copies the given Excel file into the project's safety folder and tracks it."""
        src = Path(original_file_path)
        dest = self.project_dir / "imports" / src.name
        
        # Copy and replace if exists
        if src.resolve() != dest.resolve():
            shutil.copy2(src, dest)
            
        self.excel_roster_path = str(dest)
        self.original_excel_path = str(src.absolute())
        self.save_state()
        return dest

    def import_pdf_file(self, original_file_path: str | Path) -> Path:
        """Copies the given Student PDF into the project's safety folder and tracks it."""
        src = Path(original_file_path)
        dest = self.project_dir / "imports" / src.name
        
        # Mirror the PDF inside the project for portability
        if src.resolve() != dest.resolve():
            shutil.copy2(src, dest)
            
        self.student_pdf_path = str(dest)
        self.save_state()
        return dest

    def set_answer_key(self, version: str, answers: Dict[int, List[str]]):
        self.answer_keys[version] = AnswerKey(version=version, answers=answers)
        self.save_state()

    def get_answer_key(self, version: str) -> AnswerKey | None:
        return self.answer_keys.get(version)

    def add_or_update_result(self, res: GradingResult):
        """Adds a new result or updates an existing one by page number."""
        original_count = len(self.last_results)
        self.last_results = [r for r in self.last_results if r.page_number != res.page_number]
        self.last_results.append(res)
        
        action = "Updated" if len(self.last_results) == original_count else "Added"
        self.logger.info(f"{action} result for Page {res.page_number} (Manual Fix: {res.is_manual_fix})")
        self.save_state()

    def save_state(self):
        """Dumps all project state to the JSON file with robust error handling."""
        try:
            rel_excel = ""
            if self.excel_roster_path:
                try: rel_excel = str(Path(self.excel_roster_path).relative_to(self.project_dir))
                except ValueError: rel_excel = self.excel_roster_path
                
            rel_pdf = ""
            if self.student_pdf_path:
                try: rel_pdf = str(Path(self.student_pdf_path).relative_to(self.project_dir))
                except ValueError: rel_pdf = self.student_pdf_path

            state = {
                "project_name": self.project_name,
                "question_count": self.question_count,
                "active_models": self.active_models,
                "excel_roster_path": rel_excel.replace("\\", "/"),
                "student_id_col": self.student_id_col,
                "grade_output_col": self.grade_output_col,
                "student_pdf_path": rel_pdf.replace("\\", "/"),
                "mark_sensitivity": self.mark_sensitivity,
                "answer_keys": {ver: ak.to_dict() for ver, ak in self.answer_keys.items()},
                "last_results": [r.to_dict() for r in self.last_results]
            }
            with open(self.state_file, "w", encoding="utf-8") as f:
                json.dump(state, f, indent=4)
            self.logger.debug(f"Project state saved successfully. Total results: {len(self.last_results)}")
        except Exception as e:
            self.logger.error(f"CRITICAL: Failed to save project state: {e}", exc_info=True)
            raise

    def load_state(self):
        """Loads state from the JSON file with validation."""
        try:
            if not self.state_file.exists():
                self.logger.warning("No state file found to load.")
                return

            with open(self.state_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                
            self.project_name = data.get("project_name", "Untitled")
            self.question_count = int(data.get("question_count", 60))
            self.active_models = int(data.get("active_models", 6))
            
            loaded_excel = data.get("excel_roster_path", "")
            if loaded_excel:
                p = Path(loaded_excel)
                self.excel_roster_path = str(p) if p.is_absolute() else str(self.project_dir / p)
            else:
                self.excel_roster_path = ""
                
            self.student_id_col = data.get("student_id_col", "")
            self.grade_output_col = data.get("grade_output_col", "")
            
            loaded_pdf = data.get("student_pdf_path", "")
            if loaded_pdf:
                p = Path(loaded_pdf)
                self.student_pdf_path = str(p) if p.is_absolute() else str(self.project_dir / p)
            else:
                self.student_pdf_path = ""
                
            self.mark_sensitivity = int(data.get("mark_sensitivity", 75))
            
            raw_keys = data.get("answer_keys", {})
            self.answer_keys = {
                ver: AnswerKey.from_dict(ak_data) for ver, ak_data in raw_keys.items()
            }
            
            raw_results = data.get("last_results", [])
            loaded_results = []
            for r_data in raw_results:
                try:
                    loaded_results.append(GradingResult.from_dict(r_data))
                except Exception as ex:
                    self.logger.warning(f"Skipping a corrupted result entry: {ex}")
            
            self.last_results = loaded_results
            self.logger.info(f"Project state loaded. Name: {self.project_name}, Results: {len(self.last_results)}")
        except Exception as e:
            self.logger.error(f"Failed to load project state: {e}", exc_info=True)

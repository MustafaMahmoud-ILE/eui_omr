import pandas as pd
from pathlib import Path
from typing import List, Dict

class ExcelManager:
    def __init__(self, file_path: str | Path):
        self.file_path = Path(file_path)

    def get_columns(self) -> List[str]:
        """Returns the list of columns available in the first sheet."""
        if not self.file_path.exists():
            return []
        try:
            df = pd.read_excel(self.file_path, nrows=0)
            return list(df.columns)
        except Exception:
            return []

    def export_grades(
        self,
        results_data: List[dict],
        student_id_col: str,
        grade_col: str,
        question_count: int,
        answer_keys: dict
    ):
        """
        Updates the target grade_col in the main sheet based on the given student IDs,
        and creates a new 'Breakdown' sheet.
        """
        if not self.file_path.exists():
            raise FileNotFoundError("Excel master file is missing.")
            
        # 1. Load Original Data and detect sheet name
        try:
            with pd.ExcelFile(self.file_path) as xl:
                sheet_names = xl.sheet_names
                if not sheet_names:
                    raise ValueError("Excel file has no sheets.")
                main_sheet_name = sheet_names[0]
                df_roster = pd.read_excel(xl, sheet_name=main_sheet_name)
        except Exception as e:
            raise RuntimeError(f"Could not read Excel file: {e}")

        # Ensure grade_col exists, if not create it
        if grade_col not in df_roster.columns:
            df_roster[grade_col] = None

        # 2. Build Breakdown Sheet Data
        breakdown_rows = []
        
        # Create a mapping of normalized IDs in roster for fast lookup
        # We store (normalized_id -> original_index)
        id_map = {}
        for idx, val in df_roster[student_id_col].items():
            norm = self._normalize_id(val)
            if norm:
                id_map[norm] = idx

        for res in results_data:
            sid = self._normalize_id(res["student_id"])
            version = res["version"]
            
            # Calculate Score
            score = 0
            ans_key = answer_keys.get(version)
            if ans_key and not res.get("id_error") and not res.get("version_error"):
                # Handle both AnswerKey objects and raw dicts from project state
                key_data = ans_key.answers if hasattr(ans_key, "answers") else ans_key
                if isinstance(key_data, dict):
                    for q in range(1, question_count + 1):
                        correct_answers = key_data.get(q, [])
                        student_answers = res["answers"].get(q, [])
                        
                        if any(a in correct_answers for a in student_answers):
                            score += 1
                        
            # Update original roster using normalized mapping
            if sid in id_map:
                row_idx = id_map[sid]
                df_roster.at[row_idx, grade_col] = score
                
            # Build detailed row
            d_row = {
                "Page": res["page_number"],
                "Student ID": res["student_id"],
                "Version": version,
                "Score": score,
                "ID Error": "YES" if res["id_error"] else "No",
                "Version Error": "YES" if res["version_error"] else "No",
            }
            for q in range(1, question_count + 1):
                ans_list = res["answers"].get(q, [])
                d_row[f"Q{q}"] = "".join(ans_list) if ans_list else "BLANK"
            breakdown_rows.append(d_row)
            
        df_breakdown = pd.DataFrame(breakdown_rows)
        
        # 3. Write back to Excel
        try:
            with pd.ExcelWriter(self.file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                # Use the ORIGINAL main_sheet_name to avoid duplicates
                df_roster.to_excel(writer, sheet_name=main_sheet_name, index=False)
                df_breakdown.to_excel(writer, sheet_name="OMR Breakdown", index=False)
        except PermissionError:
            raise PermissionError(f"The file '{self.file_path.name}' is open in another program. Please close it and try again.")
        except Exception as e:
            raise RuntimeError(f"Error saving to Excel: {e}")

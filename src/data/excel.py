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

    def _normalize_id(self, val) -> str:
        """
        Cleans and normalizes IDs for matching. 
        Robustly handles hyphens, leading zeros, and pandas float suffixes (.0).
        Returns a 'clean numeric string' for comparison.
        """
        if val is None: return ""
        s = str(val).strip().lower()
        if s in ["nan", "none", "null", "nat"]: return ""
        
        # Handle the common pandas float-conversion suffix '.0'
        if s.endswith(".0"):
            s = s[:-2]
            
        # Keep only digits
        s = "".join(c for c in s if c.isdigit())
        
        # Strip leading zeros
        s = s.lstrip("0")
        
        # FINAL FALLBACK: If normalization resulted in a valid integer string, 
        # that is what we return.
        if not s and any(c.isdigit() for c in str(val)):
            return "0"
            
        return s

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
            
        # 1. Load Original Data and detect dynamic sheet name
        try:
            main_sheet_name = None
            df_roster = None
            
            with pd.ExcelFile(self.file_path) as xl:
                # Search all sheets for the one containing student_id_col
                for sheet in xl.sheet_names:
                    # Read only headers to check
                    df_check = pd.read_excel(xl, sheet_name=sheet, nrows=0)
                    if student_id_col in df_check.columns:
                        main_sheet_name = sheet
                        df_roster = pd.read_excel(xl, sheet_name=sheet)
                        break
                
                if not main_sheet_name:
                    # Fallback to first sheet if not found (though UI should prevent this)
                    main_sheet_name = xl.sheet_names[0]
                    df_roster = pd.read_excel(xl, sheet_name=main_sheet_name)
                    
        except Exception as e:
            raise RuntimeError(f"Could not read Excel file: {e}")

        # Ensure grade_col exists, if not create it
        if grade_col not in df_roster.columns:
            df_roster[grade_col] = None

        # 2. Build Breakdown Sheet Data
        breakdown_rows = []
        
        # Create a mapping of normalized IDs in roster for fast lookup
        id_map = {}
        for idx, val in df_roster[student_id_col].items():
            norm = self._normalize_id(val)
            if norm:
                # Store by 'mathematical ID' (integer string) to eliminate formatting differences
                try: m_id = str(int(norm))
                except: m_id = norm
                id_map[m_id] = idx
        for res in results_data:
            # Normalize OMR ID and convert to mathematical form for comparison
            sid_norm = self._normalize_id(res["student_id"])
            try: sid_math = str(int(sid_norm))
            except: sid_math = sid_norm
            
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
                        
            # Update original roster using mathematical mapping
            if sid_math in id_map:
                row_idx = id_map[sid_math]
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
        def get_m(s):
            try: return str(int(self._normalize_id(s)))
            except: return self._normalize_id(s)
        match_count = sum(1 for res in results_data if get_m(res["student_id"]) in id_map)
        
        try:
            # 3a. Read all existing sheets to preserve them
            all_sheets = {}
            with pd.ExcelFile(self.file_path) as xl:
                for sheet in xl.sheet_names:
                    # Skip the ones we are about to overwrite/create
                    if sheet == main_sheet_name or sheet == "OMR Breakdown":
                        continue
                    all_sheets[sheet] = pd.read_excel(xl, sheet_name=sheet)
            
            # 3b. Write everything back
            with pd.ExcelWriter(self.file_path, engine="openpyxl") as writer:
                # Write back the original data first
                for sheet, df in all_sheets.items():
                    df.to_excel(writer, sheet_name=sheet, index=False)
                
                # Write updated roster and breakdown
                df_roster.to_excel(writer, sheet_name=main_sheet_name, index=False)
                df_breakdown.to_excel(writer, sheet_name="OMR Breakdown", index=False)
                
            return match_count
        except PermissionError:
            raise PermissionError(f"The file '{self.file_path.name}' is open in another program. Please close it and try again.")
        except Exception as e:
            raise RuntimeError(f"Error saving to Excel: {e}")

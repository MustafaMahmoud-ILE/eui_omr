import fitz  # PyMuPDF
import numpy as np
import traceback
from PySide6.QtCore import QThread, Signal
from pathlib import Path

from src.core.grader import OMRGrader
from src.models.schemas import GradingResult


class PDFGraderWorker(QThread):
    progress_updated = Signal(int, int)          # current, total
    review_required = Signal(object)             # Sends the GradingResult containing crops
    page_processed = Signal(object)              # Sends GradingResult if successful without review
    error_occurred = Signal(str, str)            # Title, Message
    finished = Signal()

    def __init__(self, pdf_path: str, config_path: str, expected_questions: int = 60, sensitivity: int = 75):
        super().__init__()
        self.pdf_path = pdf_path
        self.config_path = config_path
        self.expected_questions = expected_questions
        self.sensitivity = sensitivity
        self._is_cancelled = False
        self._is_paused = False

    def cancel(self):
        self._is_cancelled = True

    def resume_after_review(self):
        self._is_paused = False

    def run(self):
        try:
            grader = OMRGrader(self.config_path, sensitivity=self.sensitivity)
            doc = fitz.open(self.pdf_path)
            total_pages = len(doc)
            
            for i in range(total_pages):
                if self._is_cancelled:
                    break
                    
                page = doc.load_page(i)
                zoom = 2.0
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                
                img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                if pix.n == 3:
                    img_data = img_data[:, :, ::-1].copy()
                elif pix.n == 1:
                    img_data = np.stack((img_data,)*3, axis=-1).copy()
                    
                try:
                    res: GradingResult | None = grader.process_array(
                        img_data, 
                        page_num=i+1, 
                        expected_questions=self.expected_questions
                    )
                    
                    if res is None:
                        # Blank page detected. Skip safely.
                        self.progress_updated.emit(i + 1, total_pages)
                        continue
                        
                    # Always emit page_processed, but the UI will know if it needs review by checking the flags in res
                    self.page_processed.emit(res)
                    
                    if res.id_error or res.version_error or len(res.question_errors) > 0:
                        # We still emit this for potential UI tracking, but we don't pause.
                        self.review_required.emit(res)
                        
                except Exception as e:
                    traceback.print_exc()
                    self.error_occurred.emit("Page Error", f"Error on page {i+1}: {str(e)}")
                    
                self.progress_updated.emit(i + 1, total_pages)
                
            doc.close()
            
        except Exception as e:
            traceback.print_exc()
            self.error_occurred.emit("PDF Error", str(e))
            
        finally:
            self.finished.emit()


class AutoTuneWorker(QThread):
    finished = Signal(int) # Returns the best sensitivity
    error = Signal(str)

    def __init__(self, pdf_path: str, config_path: str, expected_questions: int):
        super().__init__()
        self.pdf_path = pdf_path
        self.config_path = config_path
        self.expected_questions = expected_questions

    def run(self):
        try:
            doc = fitz.open(self.pdf_path)
            total_pages = len(doc)
            if total_pages == 0:
                self.error.emit("PDF is empty")
                return
            
            # Sample up to 3 pages for a more 'strong' and robust analysis
            sample_images = []
            num_samples = min(3, total_pages)
            
            for i in range(num_samples):
                page = doc.load_page(i)
                zoom = 2.0
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                if pix.n == 3: 
                    img = img[:, :, ::-1].copy()
                sample_images.append(img)
            
            doc.close()

            grader = OMRGrader(self.config_path)
            # Pass the ACTUAL list of images for a truly intelligent optimization
            best_sens = grader.optimize_sensitivity(sample_images, self.expected_questions)
            self.finished.emit(best_sens)
        except Exception as e:
            traceback.print_exc()
            self.error.emit(str(e))

import sys
import os
from pathlib import Path

# Fix import path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.core.grader import OMRGrader
import pandas as pd
import fitz
import numpy as np

# --- Configuration ---
PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "config.json"
PDF_PATH = PROJECT_ROOT / "test_data_gen" / "mock_exam_20pages.pdf"
GT_EXCEL = PROJECT_ROOT / "test_data_gen" / "ground_truth_answers.xlsx"

def run_headless_verification():
    print("--- Automated OMR Verification Started ---")
    
    # 1. Load Ground Truth
    if not GT_EXCEL.exists():
        print(f"Error: {GT_EXCEL} not found.")
        return
    gt_df = pd.read_excel(GT_EXCEL)
    
    # 2. Initialize Grader & Auto-Tune
    print("Auto-tuning sensitivity...")
    grader = OMRGrader(CONFIG_PATH)
    
    # Select 5 random pages for tuning
    import random
    doc = fitz.open(str(PDF_PATH))
    num_pages = len(doc)
    num_samples = min(5, num_pages)
    indices = random.sample(range(num_pages), num_samples)
    
    sample_imgs = []
    for i in indices:
        p = doc.load_page(i)
        px = p.get_pixmap(matrix=fitz.Matrix(200/72, 200/72), alpha=False)
        im = np.frombuffer(px.samples, dtype=np.uint8).reshape(px.h, px.w, px.n)
        if px.n == 3: im = im[:, :, ::-1].copy()
        sample_imgs.append(im)
    
    best_sens = grader.optimize_sensitivity(sample_imgs, expected_questions=5)
    print(f"Optimal Sensitivity Found: {best_sens}")
    grader = OMRGrader(CONFIG_PATH, sensitivity=best_sens)
    
    # 3. Process PDF
    total_pages = len(doc)
    results = []

    print(f"Processing {total_pages} pages using OMRGrader engine (200 DPI)...")
    
    for i in range(total_pages):
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=fitz.Matrix(200/72, 200/72), alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        if pix.n == 3: img = img[:, :, ::-1].copy()
        
        res = grader.process_array(img, page_num=i+1, expected_questions=5)
        if res:
            results.append({
                "Page": res.page_number,
                "ID": res.student_id,
                "Version": res.version,
                "Q1": ",".join(res.answers.get(1, [])),
                "Q2": ",".join(res.answers.get(2, [])),
                "Q3": ",".join(res.answers.get(3, [])),
                "Q4": ",".join(res.answers.get(4, [])),
                "Q5": ",".join(res.answers.get(5, []))
            })
    doc.close()

    # 4. Compare and Score
    extracted_df = pd.DataFrame(results)
    
    correct_ids = (gt_df["Student_ID"].astype(str) == extracted_df["ID"].astype(str)).sum()
    correct_versions = (gt_df["Version"] == extracted_df["Version"]).sum()
    
    print("\n" + "="*40)
    print(f"RESULTS FOR {total_pages} PAGES:")
    print(f"- Student ID Accuracy : {correct_ids}/{total_pages} ({(correct_ids/total_pages)*100:.1f}%)")
    print(f"- Version Accuracy    : {correct_versions}/{total_pages} ({(correct_versions/total_pages)*100:.1f}%)")
    
    # Simple Question check
    total_q = total_pages * 5
    correct_q = 0
    for q in range(1, 6):
        col = f"Q{q}"
        matches = (gt_df[col].fillna("") == extracted_df[col].fillna("")).sum()
        correct_q += matches
        
    print(f"- Questions Accuracy  : {correct_q}/{total_q} ({(correct_q/total_q)*100:.1f}%)")
    print("="*40)
    print("\nVerification complete!")

if __name__ == "__main__":
    run_headless_verification()

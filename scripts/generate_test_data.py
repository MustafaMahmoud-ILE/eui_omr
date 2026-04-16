import cv2
import json
import random
import numpy as np
import pandas as pd
import fitz
from pathlib import Path

# --- Configuration ---
PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "config.json"
TEMPLATE_PATH = PROJECT_ROOT / "assets" / "template.pdf"
OUTPUT_DIR = PROJECT_ROOT / "test_data_gen"
PAGES_DIR = OUTPUT_DIR / "pages"

PAGES_DIR.mkdir(parents=True, exist_ok=True)

with open(CONFIG_PATH, "r") as f:
    CONFIG = json.load(f)

CHOICES = ["A", "B", "C", "D", "E", "F"]
VERSIONS = ["A", "B", "C", "D", "E", "F"]

def shade_bubble(img, roi_key, row, col, total_rows, total_cols):
    bbox = CONFIG[roi_key]
    cell_w = bbox["w"] / total_cols
    cell_h = bbox["h"] / total_rows
    
    # Calculate center
    cx = int(bbox["x"] + (col + 0.5) * cell_w)
    cy = int(bbox["y"] + (row + 0.5) * cell_h)
    
    # Draw a somewhat irregular circle to simulate pencil
    radius = int(min(cell_w, cell_h) * 0.35)
    
    # Add slight jitter to center
    cx += random.randint(-2, 2)
    cy += random.randint(-2, 2)
    
    # Draw 3-4 slightly offset circles to make it look "filled" by hand
    for _ in range(5):
        off_x = random.randint(-2, 2)
        off_y = random.randint(-2, 2)
        cv2.circle(img, (cx + off_x, cy + off_y), radius, (40, 40, 40), -1, cv2.LINE_AA)

def add_noise(img):
    # 1. Subtle Rotation
    h, w = img.shape[:2]
    angle = random.uniform(-0.3, 0.3)
    M = cv2.getRotationMatrix2D((w/2, h/2), angle, 1.0)
    img = cv2.warpAffine(img, M, (w, h), borderMode=cv2.BORDER_REPLICATE)
    
    # 2. Brightness/Contrast variation
    alpha = random.uniform(0.95, 1.05) 
    beta = random.randint(-10, 10)
    img = cv2.convertScaleAbs(img, alpha=alpha, beta=beta)
    
    # 3. Gaussian Noise (removed for clarity)
    # noise = np.random.normal(0, 2, img.shape).astype(np.uint8)
    # img = cv2.add(img, noise)
    
    return img

def generate_mock_session(num_students=20):
    if str(TEMPLATE_PATH).lower().endswith(".pdf"):
        doc = fitz.open(str(TEMPLATE_PATH))
        page = doc.load_page(0)
        # Higher zoom during initial render to maintain quality before our resize
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
        img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        template_raw = img_data[:, :, ::-1].copy() if pix.n == 3 else np.stack((img_data,)*3, axis=-1).copy()
        doc.close()
    else:
        template_raw = cv2.imread(str(TEMPLATE_PATH))
        
    if template_raw is None:
        print(f"Error: Could not find template at {TEMPLATE_PATH}")
        return
        
    # CRITICAL FIX: Resize template to match the resolution config.json expects!
    # Config is based on 1654x2339 (A4 @ ~200DPI)
    template = cv2.resize(template_raw, (CONFIG["a4_width"], CONFIG["a4_height"]), interpolation=cv2.INTER_AREA)

    data_log = []
    image_paths = []

    print(f"Generating {num_students} mock pages at {CONFIG['a4_width']}x{CONFIG['a4_height']}...")

    for i in range(1, num_students + 1):
        canvas = template.copy()
        
        # 1. Random Student ID
        sid_ints = [random.randint(0, 9) for _ in range(8)]
        sid_str = "".join(map(str, sid_ints))
        for col, digit in enumerate(sid_ints):
            shade_bubble(canvas, "student_id", digit, col, 10, 8)
            
        # 2. Random Version (Restrict to A-D for this test)
        v_idx = random.randint(0, 3)
        version = VERSIONS[v_idx]
        shade_bubble(canvas, "version", v_idx, 0, 6, 1)
        
        # 3. Random Answers (First 5 questions)
        ans_dict = {}
        for q in range(1, 6):
            # Pick 1 or 2 answers occasionally (for multi-choice test)
            num_ans = 1 if random.random() > 0.1 else 2
            indices = random.sample(range(4), num_ans) # typically A-D
            chosen = sorted([CHOICES[idx] for idx in indices])
            ans_dict[q] = ",".join(chosen)
            
            # Draw it
            col_name = "questions_col1" # Q1-5 are in col 1
            row_idx = q - 1
            for idx in indices:
                shade_bubble(canvas, col_name, row_idx, idx, 20, 6)
        
        # 4. Mock Signature/Name text
        cv2.putText(canvas, f"Student {i:02d}", (CONFIG["student_name"]["x"] + 20, CONFIG["student_name"]["y"] + 80),
                    cv2.FONT_HERSHEY_SCRIPT_COMPLEX, 1.5, (50, 50, 50), 2, cv2.LINE_AA)

        # 5. Apply noise and save
        canvas = add_noise(canvas)
        out_path = PAGES_DIR / f"page_{i:03d}.jpg"
        cv2.imwrite(str(out_path), canvas)
        image_paths.append(out_path)
        
        data_log.append({
            "Page": i,
            "Student_ID": sid_str,
            "Version": version,
            "Q1": ans_dict.get(1, ""),
            "Q2": ans_dict.get(2, ""),
            "Q3": ans_dict.get(3, ""),
            "Q4": ans_dict.get(4, ""),
            "Q5": ans_dict.get(5, "")
        })

    # Create Ground Truth Excel
    df = pd.DataFrame(data_log)
    gt_path = OUTPUT_DIR / "ground_truth_answers.xlsx"
    df.to_excel(gt_path, index=False)
    print(f"Ground truth saved to: {gt_path}")

    # Create Merged PDF
    pdf_path = OUTPUT_DIR / "mock_exam_20pages.pdf"
    doc = fitz.open()
    for img_p in image_paths:
        imgdoc = fitz.open(str(img_p))
        pdfbytes = imgdoc.convert_to_pdf()
        imgpdf = fitz.open("pdf", pdfbytes)
        doc.insert_pdf(imgpdf)
    doc.save(str(pdf_path))
    doc.close()
    print(f"Mock PDF saved to: {pdf_path}")
    print("\nDone! Use the PDF and Excel file in the 'test_data_gen' folder for testing.")

if __name__ == "__main__":
    generate_mock_session(20)

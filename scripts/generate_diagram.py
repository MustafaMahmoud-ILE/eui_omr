import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import os

fig, ax = plt.subplots(figsize=(10, 6))
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

box_style = "round,pad=1.5,rounding_size=1"
font_opts = {'fontsize': 11, 'fontweight': 'bold', 'color': '#0B192E', 'ha': 'center', 'va': 'center'}

def draw_box(x, y, text, color):
    ax.add_patch(FancyBboxPatch((x, y), width=28, height=12, boxstyle=box_style, 
                                facecolor=color, edgecolor='#0B192E', linewidth=1.5))
    ax.text(x + 14, y + 6, text, **font_opts)

def draw_arrow(p1, p2, text="", text_offset_x=0, text_offset_y=3):
    ax.add_patch(FancyArrowPatch(p1, p2, arrowstyle='->', mutation_scale=20, linewidth=2, color='#1A73E8'))
    if text:
        ax.text((p1[0]+p2[0])/2 + text_offset_x, (p1[1]+p2[1])/2 + text_offset_y, text, 
                fontsize=10, color='#1A73E8', fontweight='bold', ha='center', va='center')

# Coordinates (x, y) Bottom-Left of box
input_pos = (5, 80)
calib_pos = (5, 55)
grade_pos = (5, 30)
review_pos = (45, 30)
export_pos = (45, 5)

# Draw Boxes
draw_box(*input_pos, "Scan PDF &\nExcel Roster", "#DBEAFE") # Blue
draw_box(*calib_pos, "Smart Auto-Tune\nCalibration", "#D1FAE5") # Green
draw_box(*grade_pos, "Batch Peak Detection\nEngine", "#FFEDD5") # Orange
draw_box(*review_pos, "Human-in-the-Loop\nReview", "#FEE2E2") # Red
draw_box(*export_pos, "Export Final Grades\nto Excel", "#F3E8FF") # Purple

# Draw Arrows
# Input to Calibration
draw_arrow((19, 80), (19, 67))
# Calibration to Grading
draw_arrow((19, 55), (19, 42))

# Grading to Review (Right)
draw_arrow((33, 36), (45, 36), "Flag Conflicts")

# Review to Export (Down)
draw_arrow((59, 30), (59, 17), "Resolve")

# Grading to Export (L-shape bypass)
ax.plot([19, 19], [30, 11], color='#1A73E8', linewidth=2)
draw_arrow((19, 11), (45, 11), "Clean Data", 0, 2)

img_path = "user_guide/workflow_diagram.pdf"
os.makedirs("user_guide", exist_ok=True)
plt.savefig(img_path, format="pdf", bbox_inches='tight')
print(f"Saved {img_path}")

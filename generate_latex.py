"""
generate_latex.py
Extracts figures from executed notebooks into latex_report/figures/.
Run this once before compiling main.tex.
"""

import json, base64, os, sys

DATA_DIR = r'c:\Users\I764176\Documents\University\Semester 3\Data Mining'
OUT_DIR  = os.path.join(DATA_DIR, 'latex_report')
FIG_DIR  = os.path.join(OUT_DIR, 'figures')
os.makedirs(FIG_DIR, exist_ok=True)


def extract_image(nb_path, cell_idx, out_name):
    with open(nb_path, 'r', encoding='utf-8') as f:
        nb = json.load(f)
    cell = nb['cells'][cell_idx]
    for out in cell.get('outputs', []):
        data = out.get('data', {})
        if 'image/png' in data:
            raw = data['image/png']
            if isinstance(raw, list):
                raw = ''.join(raw)
            out_path = os.path.join(FIG_DIR, out_name)
            with open(out_path, 'wb') as fh:
                fh.write(base64.b64decode(raw))
            print(f'  OK  figures/{out_name}')
            return True
    print(f'  WARN  No image at cell {cell_idx} in {os.path.basename(nb_path)}')
    return False


nb = lambda name: os.path.join(DATA_DIR, 'Code', f'{name}.ipynb')

print('Extracting figures ...')
extract_image(nb('section5_evaluation_colab'),           11, 'fig03_model_comparison.png')
extract_image(nb('section6_threshold_experiment_colab'), 17, 'fig04_threshold_cms.png')
extract_image(nb('section4b_random_forest_colab'),       13, 'fig05_rf_cm.png')
extract_image(nb('section4b_random_forest_colab'),       15, 'fig06_rf_importance.png')
extract_image(nb('section7_error_analysis_colab'),        8, 'fig15a_error_flow_cm.png')
extract_image(nb('section7_error_analysis_colab'),       10, 'fig15b_confidence.png')
extract_image(nb('section7_error_analysis_colab'),       16, 'fig15c_feat_separation.png')
extract_image(nb('section8_feature_engineering_colab'),   7, 'fig17a_cv_folds.png')
extract_image(nb('section8_feature_engineering_colab'),  13, 'fig17b_cm_comparison.png')
extract_image(nb('section8_feature_engineering_colab'),  15, 'fig17c_feat_importance.png')

print(f'\nDone. Upload the latex_report/ folder to Overleaf, or compile locally:')
print(f'  cd {OUT_DIR}')
print(f'  pdflatex main.tex && pdflatex main.tex')

# Steam Games – Predicting Game Success
### IE500 Data Mining | University of Mannheim | BSS 2026
**Team 9 – Brewed Clusters**

Satwik Kumar Lohani · Arpitha Shivaprasad Kori · Zubair Ashfaq · Beyza Ünlü · Muhammad Sameer Siddiqui · Esha Raheel

---

## Project Overview

This project applies data mining and machine learning to a dataset of ~122,000 Steam games. The goal is to **predict whether a game will be well-received, mixed, or poorly-received** (multiclass classification: Good / Mixed / Bad) based on observable metadata such as genre, price, tags, and platform availability.

Beyond prediction, we use SHAP values and feature importances to identify which game attributes most strongly drive positive reception.

---

## Repository Structure

```
steam-games-data-mining/
│
├── Datasets/                               # Cleaned, model-ready data
│   ├── train_multiclass.csv                # 80% split — for model training (45,324 × 148)
│   └── test_multiclass.csv                 # 20% split — for model evaluation (11,331 × 148)
│
├── Code/
│   ├── section1_business_understanding.ipynb
│   ├── section2_preprocessing.ipynb
│   ├── run_preprocessing.py
│   ├── section3_eda.ipynb
│   ├── section4a_logistic.ipynb
│   ├── section4b_random_forest.ipynb
│   ├── section4c_xgboost.ipynb
│   └── section5_evaluation.ipynb
│
└── ProjectOutline - Brewed Clusters.pdf
```

---

## Getting Started

### 1. Clone the repository
```bash
git clone https://github.com/lohaniSatwik/steam-games-data-mining.git
cd steam-games-data-mining
```

### 2. Set up the Python environment
```bash
cd Code
python -m venv .venv

# Windows
.venv\Scripts\activate

# Mac / Linux
source .venv/bin/activate

pip install pandas numpy matplotlib seaborn scikit-learn xgboost
```

### 3. Load the data in your notebook

All datasets are in the `Datasets/` folder at the root of the repo.

**For Modelling** (pre-split):
```python
import pandas as pd

train = pd.read_csv('../Datasets/train_multiclass.csv')
test  = pd.read_csv('../Datasets/test_multiclass.csv')

X_train, y_train = train.drop(columns=['label_multiclass']), train['label_multiclass']
X_test,  y_test  = test.drop(columns=['label_multiclass']),  test['label_multiclass']
```

---

## About the Dataset

| Property | Value |
|----------|-------|
| Total games | 56,655 (filtered to ≥10 reviews) |
| Train set | 45,324 rows (80%) |
| Test set | 11,331 rows (20%) |
| Features | 147 (numeric + binary dummies) |
| Label | `label_multiclass` — Good / Mixed / Bad |
| Class balance (train) | Good: 63.3% / Mixed: 28.3% / Bad: 8.4% |
| Missing values | None (median imputed) |
| Scaling | Not applied — scale inside your CV pipeline |

**Feature groups:**
- 12 numeric columns (price, achievements, playtime, etc.)
- 28 genre dummy columns (`genre_*`)
- 58 category dummy columns (`cat_*`)
- 50 tag dummy columns (`tag_*`) — top 50 by frequency
- 3 release era columns (`era_*`)

> **Note on potential leakage:** `Average playtime forever`, `Median playtime forever`, and `Recommendations` are post-release metrics. They are included for now — consider running your model with and without them to compare the impact.

---

## Team Responsibilities

| Section | Owner | Notebook |
|---------|-------|----------|
| Preprocessing | Satwik Kumar Lohani | `section2_preprocessing.ipynb` |
| EDA | Arpitha Shivaprasad Kori | `section3_eda.ipynb` |
| Logistic Regression | Zubair Ashfaq | `section4a_logistic.ipynb` |
| Random Forest | Beyza Ünlü | `section4b_random_forest.ipynb` |
| XGBoost | Muhammad Sameer Siddiqui | `section4c_xgboost.ipynb` |
| Evaluation, SHAP & Report | Esha Raheel | `section5_evaluation.ipynb` |

---

## Modelling Notes

- **Task:** Multiclass classification — Good / Mixed / Bad
- **Evaluation metric:** Macro F1-score (primary), Precision and Recall per class (secondary)
- **Cross-validation:** 10-fold stratified outer loop; 3-fold inner loop for hyperparameter tuning
- **Class imbalance:** Use `class_weight='balanced'` inside your model — do NOT oversample the training data before CV
- **Scaling:** Apply `StandardScaler` inside each model's `sklearn.Pipeline` on the continuous columns: `log_price`, `Required age`, `DiscountDLC count`, `Achievements`, `Average playtime forever`, `Median playtime forever`, `Recommendations`, `Metacritic score`, `n_languages`
- **Test set:** Only evaluate on `test_multiclass.csv` once, at the very end

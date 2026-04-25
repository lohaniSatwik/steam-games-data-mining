# Steam Games – Predicting Game Success
### IE500 Data Mining | University of Mannheim | BSS 2026
**Team 9 – Brewed Clusters**

Satwik Kumar Lohani · Arpitha Shivaprasad Kori · Zubair Ashfaq · Beyza Ünlü · Muhammad Sameer Siddiqui · Esha Raheel

---

## Project Overview

This project applies data mining and machine learning to a dataset of ~122,000 Steam games. The goal is to **predict whether a game will receive a positive or negative review outcome** (binary classification: Good / Bad) based on its observable metadata attributes such as genre, price, tags, and platform availability.

Beyond prediction, we use SHAP values and feature importances to identify which game attributes most strongly drive positive reception.

---

## Repository Structure

```
steam-games-data-mining/
│
├── Code/
│   ├── Data/
│   │   └── processed/                  # Generated locally — do not commit
│   │
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

### 2. Download the raw dataset
The raw `games.csv` file (372MB) is not stored in this repository. Download it from Kaggle:

> **[Steam Games Dataset – FronkonGames](https://www.kaggle.com/fronkongames/steam-games-dataset)**

Place the downloaded file at:
```
Code/Data/games.csv
```

### 3. Set up the Python environment
```bash
cd Code
python -m venv .venv

# Windows
.venv\Scripts\activate

# Mac / Linux
source .venv/bin/activate

pip install pandas numpy matplotlib seaborn scikit-learn pyarrow joblib
```

### 4. Run preprocessing
```bash
cd Code
python run_preprocessing.py
```

This generates the three files in `Code/Data/processed/` that all modelling notebooks depend on.
**You only need to run this once.**

---

## Team Responsibilities

| Section | Owner | Notebook |
|---------|-------|----------|
| Preprocessing | Satwik Kumar Lohani | `section2_preprocessing.ipynb` |
| EDA | Arpitha Shivaprasad Kori | `section3_eda.ipynb` |
| Baseline + Logistic Regression | Zubair Ashfaq | `section4a_logistic.ipynb` |
| Random Forest | Beyza Ünlü | `section4b_random_forest.ipynb` |
| XGBoost + Hyperparameter Tuning | Muhammad Sameer Siddiqui | `section4c_xgboost.ipynb` |
| Evaluation, SHAP & Report | Esha Raheel | `section5_evaluation.ipynb` |

---

## Modelling Notes

- **Evaluation metric:** Macro F1-score (primary), AUC-ROC and Precision-Recall AUC (secondary)
- **Cross-validation:** 10-fold stratified CV (outer loop); 3-fold inner loop for hyperparameter tuning
- **Class imbalance:** 71.4% Good / 28.6% Bad — use `class_weight='balanced'` and SMOTE
- **Scaling:** Apply `StandardScaler` inside each model's `sklearn.Pipeline` on the continuous columns (log_price, Required age, DiscountDLC count, Achievements, Average playtime forever, Median playtime forever, Recommendations, Metacritic score, n_languages)

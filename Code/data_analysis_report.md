# Data Analysis Report -- Steam Games Dataset
**Team 9 - Brewed Clusters | IE500 Data Mining | University of Mannheim**  
**Date:** 2026-04-02  
**Dataset:** Steam Games (FronkonGames, Kaggle) -- `Data/games.csv`

---

## 1. Dataset Overview

| Property | Value |
|---|---|
| Total rows (games) | 122,611 |
| Total columns (attributes) | 39 |
| File format | CSV |
| Source | Kaggle -- FronkonGames Steam Games Dataset (updated 2024) |

---

## 2. Critical Bug: CSV Parsing Error

### What we found
When loading the file with the default `pd.read_csv()` call, pandas silently misreads the entire dataset. Every column from `About the game` onwards is shifted one position to the right.

**Root cause:** The `Supported languages` field contains commas inside quoted strings (e.g. `"['English', 'French']"`). This produces **40 fields per data row** against a **39-column header**, causing a one-column offset for the remainder of every row.

**Symptom that revealed the bug:** `Peak CCU` appeared to have a maximum value of 21 and 99% zeros -- impossible for a real-world dataset containing Counter-Strike 2 and Dota 2.

### Fix
Read the CSV using Python's built-in `csv.reader` and manually merge the split column before constructing the DataFrame. All notebooks must use this corrected loading method.

```python
import csv, pandas as pd

rows_data = []
with open('Data/games.csv', 'r', encoding='utf-8') as f:
    reader = csv.reader(f)
    header = next(reader)
    for row in reader:
        if len(row) == 40:
            # merge the two halves of the split 'Supported languages' cell
            row = row[:9] + [row[9] + ',' + row[10]] + row[11:]
        rows_data.append(row)

df = pd.DataFrame(rows_data, columns=header)
```

### Impact
Without this fix, every column from index 9 onwards is wrong. This affects:
- `Positive`, `Negative` (review counts)
- `Average playtime forever`
- `Recommendations`
- `Genres`, `Tags`
- Every other column used for modelling

**This fix must be applied in every notebook in the project.**

---

## 3. Missing Values

| Column | Missing Count | Missing % | Decision |
|---|---|---|---|
| Movies | 122,611 | 100.0% | Drop -- entirely empty |
| Score rank | 122,571 | 99.97% | Drop -- nearly empty |
| Metacritic url | 118,355 | 96.5% | Drop -- not useful |
| Reviews (text) | 110,541 | 90.2% | Drop -- numeric counts available instead |
| Notes | 100,153 | 81.7% | Drop -- editorial notes, not useful |
| Website | 72,935 | 59.5% | Drop -- not a predictive feature |
| Support url | 68,469 | 55.8% | Drop -- not a predictive feature |
| Tags | 39,265 | 32.0% | Keep -- treat missing as "no tags" |
| Support email | 22,263 | 18.2% | Drop -- not a predictive feature |
| Categories | 8,953 | 7.3% | Keep -- fill missing with empty string |
| Publishers | 8,909 | 7.3% | Keep -- fill missing with "Unknown" |
| About the game | 8,449 | 6.9% | Drop -- free text, not used in models |
| Developers | 8,437 | 6.9% | Keep -- fill missing with "Unknown" |
| Genres | 8,413 | 6.9% | Keep -- treat missing as "no genre" |
| Screenshots | 6,018 | 4.9% | Drop -- URLs, not useful |
| Header image | 81 | 0.07% | Drop -- URL, not useful |

**Columns we need for modelling have very low missing rates** (Genres 6.9%, Tags 32%) -- both fixable.

---

## 4. Task 1 -- Classification Target Analysis

### Target derivation
The classification label is derived from review counts:

```
review_ratio = Positive / (Positive + Negative)
label = "Good" if review_ratio >= 0.70 else "Bad"
```

### Filtering
Games with **zero reviews are removed** -- we cannot label a game as Good or Bad if nobody reviewed it.

| Group | Count | Share |
|---|---|---|
| Zero-review games (removed) | 39,662 | 32.3% |
| Games with at least 1 review (kept) | 82,949 | 67.7% |

### Class distribution (after filtering)

| Class | Count | Share |
|---|---|---|
| Good (ratio >= 0.70) | 56,803 | 68.5% |
| Bad (ratio < 0.70) | 26,146 | 31.5% |
| **Imbalance ratio** | **2.2 : 1** | |

### Assessment
The 2.2:1 imbalance is **manageable** -- not extreme. The plan already accounts for this using SMOTE oversampling and `class_weight='balanced'`. No change to the original plan is needed.

---

## 5. Task 2 -- Regression Target Analysis

### Original target: Peak CCU
After applying the CSV fix, the corrected Peak CCU values are:

| Statistic | Value |
|---|---|
| Min | 0 |
| Median | 0 |
| 95th percentile | 9 |
| Max | 1,013,936 (Counter-Strike 2) |
| Zero values | 102,935 (84.0%) |
| Skewness | 210 |

**Problem:** 84% of games have a Peak CCU of zero. A regression model trained on this data would simply learn to predict "near zero" for almost everything and still score well -- making it useless in practice.

### Alternative target comparison

| Column | Non-zero % | Skewness | Notes |
|---|---|---|---|
| Peak CCU | 16% | 210 | 84% zeros -- original choice |
| Recommendations | 17% | 114 | Similar sparsity problem |
| Average playtime | 21% | 263 | 79% zeros -- worse than CCU |
| Positive reviews | 65% | 178 | Cannot use -- already used to build classification label (circular) |
| **Estimated owners** | **82%** | **72** | **Best candidate** |

### Decision: Replace Peak CCU with Estimated owners

`Estimated owners` is the best regression target because:
- **82% of games have a non-zero value** -- the model has meaningful data to learn from
- **Skewness of 72** -- still needs log-transform, but much less extreme than CCU's 210
- **Business relevance** -- predicting how many people bought a game is arguably more useful than peak concurrent players
- **Correct top games** -- CS2, PUBG, Dota 2, Black Myth: Wukong appear at the top, confirming data quality

### Estimated owners distribution (after CSV fix)

| Ownership bucket | Games | Share |
|---|---|---|
| 0 | 21,641 | 17.7% |
| 1 - 50,000 | 86,800 | 70.8% |
| 50,000 - 100,000 | 5,355 | 4.4% |
| 100,000 - 500,000 | 6,307 | 5.1% |
| 500,000 - 1,000,000 | 1,154 | 0.9% |
| 1,000,000 - 5,000,000 | 1,134 | 0.9% |
| 5,000,000+ | 220 | 0.2% |

### Preprocessing for Estimated owners
The raw column stores range strings (e.g. `"50000 - 100000"`). We convert to a numeric midpoint:

```python
def parse_owners(s):
    parts = str(s).split(' - ')
    return (int(parts[0]) + int(parts[1])) / 2
```

A log-transform is then applied before training, identical to what was planned for Peak CCU.

**Limitation to note in the report:** Midpoint conversion introduces imprecision (e.g. a game with 51,000 owners and one with 99,000 both map to 75,000). This is standard practice for this dataset.

### Scope of change
Switching the regression target from Peak CCU to Estimated owners affects **only Task 2**. Classification and clustering are unchanged.

---

## 6. Task 3 -- Clustering Data Readiness

### Genre distribution (top 10)

| Genre | Count |
|---|---|
| Indie | 80,630 |
| Casual | 50,210 |
| Action | 46,220 |
| Adventure | 45,141 |
| Simulation | 24,114 |
| Strategy | 22,400 |
| RPG | 20,972 |
| Free To Play | 12,172 |
| Early Access | 11,091 |
| Sports | 4,882 |

### Top tags (top 10)

| Tag | Count |
|---|---|
| Singleplayer | 50,350 |
| Indie | 48,552 |
| Action | 36,809 |
| Casual | 36,564 |
| Adventure | 35,196 |
| 2D | 26,792 |
| Simulation | 17,893 |
| Strategy | 17,848 |
| Puzzle | 16,368 |
| Atmospheric | 16,281 |

### Assessment
The genre and tag distributions show clear natural groupings -- strong signal for clustering. Expected segments include free-to-play multiplayer games, premium indie titles, casual puzzle games, and AAA action/RPG titles. Data is suitable for clustering as planned.

---

## 7. Platform Coverage

| Platform | Games | Share |
|---|---|---|
| Windows | 122,567 | 100.0% |
| Mac | 21,292 | 17.4% |
| Linux | 15,706 | 12.8% |

Near-universal Windows coverage. Mac and Linux availability can be used as binary features in models.

---

## 8. Summary of Changes to the Project Plan

| Item | Original Plan | Updated Decision | Reason |
|---|---|---|---|
| CSV loading | `pd.read_csv()` default | `csv.reader` with column merge fix | Parsing bug -- 40 cols vs 39-col header |
| Regression target | `Peak CCU` | `Estimated owners (midpoint)` | Peak CCU has 84% zeros -- model would be untrainable |
| Log-transform | Applied to Peak CCU | Applied to owners midpoint | Same technique, different column |
| Classification | Unchanged | Unchanged | -- |
| Clustering | Unchanged | Unchanged | -- |
| All other models | Unchanged | Unchanged | -- |

---

## 9. Next Steps

1. **Section 2 -- Preprocessing:** Implement the CSV fix, parse owner ranges, drop unused columns, handle missing values, encode genres/tags as binary columns, apply log-transform to owners, StandardScaler normalisation, export cleaned data to Parquet
2. **Section 3 -- EDA:** Visualise distributions of key features, review ratio distribution, owner distribution by genre
3. **Section 4 -- Classification models:** Baseline -> Logistic Regression -> Random Forest -> XGBoost
4. **Section 5 -- Regression models:** Baseline -> Ridge/Lasso -> Gradient Boosting
5. **Section 6 -- Clustering:** K-Means, DBSCAN, PCA visualisation

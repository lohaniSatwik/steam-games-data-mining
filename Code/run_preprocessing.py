import csv
import warnings
warnings.filterwarnings('ignore')

import pandas as pd
import numpy as np
import os
from collections import Counter
from sklearn.preprocessing import MultiLabelBinarizer

pd.set_option('display.max_columns', 50)

# ── Step 1: Load with corrected CSV routine ────────────────────────────────────
# WHY: pd.read_csv() misreads the 'Supported languages' column because it stores
# values like ['English', 'French'] — a comma-separated list inside quotes.
# This causes every column from index 9 onward to shift by one position silently.
# Python's csv.reader correctly handles quoted fields and avoids this bug.
DATA_PATH = 'Data/games.csv'
print('Step 1: Loading data...')

rows_data = []
with open(DATA_PATH, 'r', encoding='utf-8') as f:
    reader = csv.reader(f)
    header = next(reader)
    for row in reader:
        # A correctly parsed row has exactly 39 fields.
        # If we get 40, the Supported languages field was split — merge it back.
        if len(row) == 40:
            row = row[:9] + [row[9] + ',' + row[10]] + row[11:]
        rows_data.append(row)

df = pd.DataFrame(rows_data, columns=header)
print(f'  Loaded: {df.shape[0]:,} rows x {df.shape[1]} columns')

# All columns arrive as strings after csv.reader — convert the ones we need to numbers
for col in ['Peak CCU', 'Positive', 'Negative', 'Price']:
    df[col] = pd.to_numeric(df[col], errors='coerce')

print(f'  Sanity check — Peak CCU max: {df["Peak CCU"].max():,.0f}  (should be millions for CS2/Dota2)')

# ── Step 2: Drop high-missingness and non-predictive columns ──────────────────
# WHY: Columns that are mostly empty or contain no useful signal (e.g. image URLs,
# admin emails) would just add noise. We drop them before doing anything else.
print('\nStep 2: Dropping high-missingness and non-predictive columns...')
cols_to_drop = [
    'Movies',           # 100% missing
    'Score rank',       # 99.97% missing
    'Metacritic url',   # 96.5% missing
    'Reviews',          # 90% missing — raw text snippet, not structured
    'Notes',            # 81.7% missing
    'Website',          # 59.5% missing — not predictive
    'Support url',      # 55.8% missing — not predictive
    'Support email',    # not predictive
    'Header image',     # URL only — not predictive
    'Screenshots',      # URL only — not predictive
    'AppID',            # identifier, not a feature
]
cols_to_drop = [c for c in cols_to_drop if c in df.columns]
df.drop(columns=cols_to_drop, inplace=True)
print(f'  Columns remaining: {df.shape[1]}')

# ── Step 3: Filter games with >= 10 reviews ───────────────────────────────────
# WHY: Steam itself does not assign a reception label (e.g. "Mostly Positive")
# until a game has at least 10 reviews. A ratio from 1-9 reviews is statistically
# unreliable — 1 positive out of 1 review gives a 100% ratio, which is meaningless.
print('\nStep 3: Filtering games with >= 10 reviews...')
total_before = len(df)
df['total_reviews'] = df['Positive'] + df['Negative']
df = df[df['total_reviews'] >= 10].reset_index(drop=True)
total_after = len(df)
print(f'  Before : {total_before:,}')
print(f'  After  : {total_after:,}')
print(f'  Removed: {total_before - total_after:,} games with unreliable labels')

# ── Step 4: Derive review_ratio and create the classification label ───────────
# WHY: The raw data has separate Positive and Negative counts. We combine them
# into a single ratio (0.0 to 1.0), then binarise: >= 0.70 = Good, else Bad.
# 0.70 aligns with Steam's own "Mostly Positive" threshold.
# IMPORTANT: review_ratio and Positive/Negative are NOT added to the feature
# matrix X — they directly determine the label, so including them would be
# perfect data leakage (the model would just read the answer off the feature).
print('\nStep 4: Deriving review_ratio and label...')
df['review_ratio'] = df['Positive'] / df['total_reviews']
df['label'] = (df['review_ratio'] >= 0.70).astype(int)  # 1 = Good, 0 = Bad

label_counts = df['label'].value_counts()
label_pct    = df['label'].value_counts(normalize=True) * 100
print(f'  Good (1): {label_counts[1]:,} ({label_pct[1]:.1f}%)')
print(f'  Bad  (0): {label_counts[0]:,} ({label_pct[0]:.1f}%)')

# ── Step 5: Parse multi-value fields (Genres, Tags, Categories) ───────────────
# WHY: These columns store multiple values as a single comma-separated string,
# e.g. "Action,Indie,RPG". A model can't use a string — it needs numbers.
# MultiLabelBinarizer converts each unique value into its own 0/1 column:
#   "Action,Indie" → action=1, indie=1, rpg=0, ...
# For Tags we only keep the top 50 most frequent to avoid a huge sparse matrix.
print('\nStep 5: Parsing Genres, Categories, Tags...')

def parse_list_column(series):
    """Split a comma-separated string into a clean list of values."""
    return series.fillna('').apply(
        lambda x: [v.strip() for v in x.split(',') if v.strip()]
    )

# Genres
genres_lists   = parse_list_column(df['Genres'])
mlb_genres     = MultiLabelBinarizer()
genres_dummies = pd.DataFrame(
    mlb_genres.fit_transform(genres_lists),
    columns=[f'genre_{g}' for g in mlb_genres.classes_],
    dtype=int
)

# Categories
cats_lists   = parse_list_column(df['Categories'])
mlb_cats     = MultiLabelBinarizer()
cats_dummies = pd.DataFrame(
    mlb_cats.fit_transform(cats_lists),
    columns=[f'cat_{c}' for c in mlb_cats.classes_],
    dtype=int
)

# Tags — top 50 only
tags_lists = parse_list_column(df['Tags'])
tag_counts = Counter(tag for tags in tags_lists for tag in tags)
top50_tags = {tag for tag, _ in tag_counts.most_common(50)}
tags_filtered = tags_lists.apply(lambda tags: [t for t in tags if t in top50_tags])
mlb_tags      = MultiLabelBinarizer(classes=sorted(top50_tags))
tags_dummies  = pd.DataFrame(
    mlb_tags.fit_transform(tags_filtered),
    columns=[f'tag_{t}' for t in mlb_tags.classes_],
    dtype=int
)

print(f'  Genre columns    : {genres_dummies.shape[1]}')
print(f'  Category columns : {cats_dummies.shape[1]}')
print(f'  Tag columns      : {tags_dummies.shape[1]}')

# ── Step 6: Feature engineering ───────────────────────────────────────────────
print('\nStep 6: Feature engineering...')

# log_price: price is heavily skewed (many free games, a few expensive ones).
# Log-transform compresses the scale so a $60 game isn't 60x more impactful
# than a $1 game in the model's eyes.
df['Price'] = pd.to_numeric(df['Price'], errors='coerce').fillna(0)
df['log_price'] = np.log1p(df['Price'])   # log(1 + price) avoids log(0) for free games

# release_era: the raw release date is too granular. We group years into
# three buckets so the model learns era-level patterns, not specific years.
df['release_year'] = pd.to_datetime(df['Release date'], errors='coerce').dt.year

def year_to_era(y):
    if pd.isna(y):  return 'Unknown'
    if y < 2015:    return 'pre-2015'
    if y <= 2019:   return '2015-2019'
    return '2020+'

df['release_era'] = df['release_year'].apply(year_to_era)
# get_dummies turns the era string into separate 0/1 columns, one per era bucket
df = pd.get_dummies(df, columns=['release_era'], prefix='era')
# Ensure era columns are integer (pandas 2+ returns bool by default)
era_cols = [c for c in df.columns if c.startswith('era_')]
df[era_cols] = df[era_cols].astype(int)

# n_languages: instead of storing the full language list, we just count how
# many languages a game supports — more languages = wider audience reach.
def count_languages(s):
    if pd.isna(s) or str(s).strip() in ('', '[]'):
        return 0
    return len([v for v in str(s).split(',') if v.strip()])

df['n_languages'] = df['Supported languages'].apply(count_languages)

# Platform flags: convert string 'True'/'False' to integer 1/0
for col in ['Windows', 'Mac', 'Linux']:
    df[col] = df[col].map({'True': 1, 'False': 0, True: 1, False: 0}).fillna(0).astype(int)

# Convert remaining numeric columns from string to float
numeric_cols = [
    'Required age', 'DiscountDLC count', 'Achievements',
    'Average playtime forever', 'Median playtime forever',
    'Recommendations', 'Metacritic score'
]
for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

# Metacritic score: a value of 0 means the game has NO Metacritic entry,
# not that it scored 0. Keeping 0 would mislead the model. Replace with NaN
# so it gets filled with the median of games that DO have a score.
df['Metacritic score'] = df['Metacritic score'].replace(0, np.nan)

print('  Done.')

# ── Step 7: Assemble the feature matrix ───────────────────────────────────────
# We combine all feature groups into one clean DataFrame X.
# Deliberately excluded:
#   - Positive, Negative, review_ratio, total_reviews → these determine the label (leakage)
#   - Name, Developer, Publisher → raw text with no numeric encoding
#   - Estimated owners → post-release metric (leakage concern, deferred)
print('\nStep 7: Assembling feature matrix...')

base_numeric = [
    'log_price',
    'Required age',
    'DiscountDLC count',
    'Achievements',
    'Average playtime forever',   # NOTE: potential post-release leakage — may be removed later
    'Median playtime forever',    # NOTE: potential post-release leakage — may be removed later
    'Recommendations',            # NOTE: potential post-release leakage — may be removed later
    'Metacritic score',
    'Windows', 'Mac', 'Linux',
    'n_languages',
]
feature_cols = base_numeric + era_cols

# Reset indices before concat — if indices don't match, rows will misalign
df_reset       = df.reset_index(drop=True)
genres_dummies = genres_dummies.reset_index(drop=True)
cats_dummies   = cats_dummies.reset_index(drop=True)
tags_dummies   = tags_dummies.reset_index(drop=True)

X = pd.concat([df_reset[feature_cols], genres_dummies, cats_dummies, tags_dummies], axis=1)
y = df_reset['label']

print(f'  Feature matrix : {X.shape}')
print(f'  Labels         : {y.shape}')

# ── Step 8: Impute remaining missing values ────────────────────────────────────
# WHY: Most ML models cannot handle NaN values at all — they throw an error.
# We fill missing values with the column median (middle value), which is more
# robust than the mean for skewed columns.
# NOTE: For maximum correctness this should happen inside CV folds, but the
# impact here is negligible given the dataset size.
print('\nStep 8: Imputing missing values...')
missing_before = X.isnull().sum().sum()
for col in X.select_dtypes(include='number').columns:
    X[col] = X[col].fillna(X[col].median())
missing_after = X.isnull().sum().sum()
print(f'  Missing before : {missing_before:,}')
print(f'  Missing after  : {missing_after:,}')

# ── Step 9: Export to CSV ─────────────────────────────────────────────────────
print('\nStep 9: Exporting to CSV...')
os.makedirs('Data/processed', exist_ok=True)

df_out = X.copy()
df_out['label'] = y.values
df_out.to_csv('Data/processed/games_processed.csv', index=False)

print(f'  Saved : Data/processed/games_processed.csv  {df_out.shape}')

print('\n' + '='*55)
print('  PREPROCESSING COMPLETE')
print('='*55)
print(f'  Final dataset size : {df_out.shape[0]:,} games')
print(f'  Feature count      : {X.shape[1]}')
print(f'  Good (1)           : {y.sum():,} ({y.mean()*100:.1f}%)')
print(f'  Bad  (0)           : {(y==0).sum():,} ({(1-y.mean())*100:.1f}%)')
print(f'  Features are UNSCALED — apply StandardScaler inside CV pipelines')
print('='*55)

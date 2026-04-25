import csv
import warnings
warnings.filterwarnings('ignore')

import pandas as pd
import numpy as np
import os
from collections import Counter
from sklearn.preprocessing import MultiLabelBinarizer, StandardScaler

pd.set_option('display.max_columns', 50)

# ── Step 1: Load with corrected CSV routine ────────────────────────────────────
DATA_PATH = 'Data/games.csv'
print('Step 1: Loading data...')

rows_data = []
with open(DATA_PATH, 'r', encoding='utf-8') as f:
    reader = csv.reader(f)
    header = next(reader)
    for row in reader:
        if len(row) == 40:
            row = row[:9] + [row[9] + ',' + row[10]] + row[11:]
        rows_data.append(row)

df = pd.DataFrame(rows_data, columns=header)
print(f'  Loaded: {df.shape[0]:,} rows x {df.shape[1]} columns')

# Convert key numeric columns
for col in ['Peak CCU', 'Positive', 'Negative', 'Price']:
    df[col] = pd.to_numeric(df[col], errors='coerce')

print(f'  Sanity check — Peak CCU max: {df["Peak CCU"].max():,.0f}')

# ── Step 2: Drop high-missingness columns ─────────────────────────────────────
print('\nStep 2: Dropping high-missingness columns...')
cols_to_drop = [
    'Movies', 'Score rank', 'Metacritic url', 'Reviews', 'Notes',
    'Website', 'Support url', 'Support email', 'Header image',
    'Screenshots', 'AppID'
]
cols_to_drop = [c for c in cols_to_drop if c in df.columns]
df.drop(columns=cols_to_drop, inplace=True)
print(f'  Columns remaining: {df.shape[1]}')

# ── Step 3: Filter games with >= 10 reviews ───────────────────────────────────
print('\nStep 3: Filtering games with >= 10 reviews...')
total_before = len(df)
df['total_reviews'] = df['Positive'] + df['Negative']
df = df[df['total_reviews'] >= 10].reset_index(drop=True)
total_after = len(df)
print(f'  Before : {total_before:,}')
print(f'  After  : {total_after:,}')
print(f'  Removed: {total_before - total_after:,} games')

# ── Step 4: Parse Estimated owners ────────────────────────────────────────────
print('\nStep 4: Parsing Estimated owners...')
def parse_owners_midpoint(s):
    try:
        parts = str(s).replace(',', '').split('-')
        low, high = float(parts[0].strip()), float(parts[1].strip())
        return (low + high) / 2
    except:
        return np.nan

df['owners_midpoint'] = df['Estimated owners'].apply(parse_owners_midpoint)
print(f'  owners_midpoint — non-null: {df["owners_midpoint"].notna().sum():,}')

# ── Step 5: Derive review_ratio and label ────────────────────────────────────
print('\nStep 5: Deriving review_ratio and label...')
df['review_ratio'] = df['Positive'] / df['total_reviews']
df['label'] = (df['review_ratio'] >= 0.70).astype(int)

label_counts = df['label'].value_counts()
label_pct    = df['label'].value_counts(normalize=True) * 100
print(f'  Good (1): {label_counts[1]:,} ({label_pct[1]:.1f}%)')
print(f'  Bad  (0): {label_counts[0]:,} ({label_pct[0]:.1f}%)')

# ── Step 6: Parse multi-value fields ─────────────────────────────────────────
print('\nStep 6: Parsing Genres, Categories, Tags...')
def parse_list_column(series):
    return series.fillna('').apply(
        lambda x: [v.strip() for v in x.split(',') if v.strip()]
    )

genres_lists = parse_list_column(df['Genres'])
mlb_genres   = MultiLabelBinarizer()
genres_dummies = pd.DataFrame(
    mlb_genres.fit_transform(genres_lists),
    columns=[f'genre_{g}' for g in mlb_genres.classes_]
)

cats_lists   = parse_list_column(df['Categories'])
mlb_cats     = MultiLabelBinarizer()
cats_dummies = pd.DataFrame(
    mlb_cats.fit_transform(cats_lists),
    columns=[f'cat_{c}' for c in mlb_cats.classes_]
)

tags_lists = parse_list_column(df['Tags'])
tag_counts = Counter(tag for tags in tags_lists for tag in tags)
top50_tags = {tag for tag, _ in tag_counts.most_common(50)}
tags_filtered = tags_lists.apply(lambda tags: [t for t in tags if t in top50_tags])
mlb_tags      = MultiLabelBinarizer(classes=sorted(top50_tags))
tags_dummies  = pd.DataFrame(
    mlb_tags.fit_transform(tags_filtered),
    columns=[f'tag_{t}' for t in mlb_tags.classes_]
)

print(f'  Genre columns    : {genres_dummies.shape[1]}')
print(f'  Category columns : {cats_dummies.shape[1]}')
print(f'  Tag columns      : {tags_dummies.shape[1]}')

# ── Step 7: Feature engineering ──────────────────────────────────────────────
print('\nStep 7: Feature engineering...')
df['Price'] = pd.to_numeric(df['Price'], errors='coerce').fillna(0)
df['log_price'] = np.log1p(df['Price'])

df['release_year'] = pd.to_datetime(df['Release date'], errors='coerce').dt.year
def year_to_era(y):
    if pd.isna(y):  return 'Unknown'
    if y < 2015:    return 'pre-2015'
    if y <= 2019:   return '2015-2019'
    return '2020+'

df['release_era'] = df['release_year'].apply(year_to_era)
df = pd.get_dummies(df, columns=['release_era'], prefix='era')

def count_languages(s):
    if pd.isna(s) or str(s).strip() in ('', '[]'): return 0
    return len([v for v in str(s).split(',') if v.strip()])

df['n_languages'] = df['Supported languages'].apply(count_languages)

for col in ['Windows', 'Mac', 'Linux']:
    df[col] = df[col].map({'True': 1, 'False': 0, True: 1, False: 0}).fillna(0).astype(int)

numeric_cols = [
    'Required age', 'DiscountDLC count', 'Achievements',
    'Average playtime forever', 'Median playtime forever',
    'Recommendations', 'Metacritic score'
]
for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

print('  Done.')

# ── Step 8: Assemble feature matrix ──────────────────────────────────────────
print('\nStep 8: Assembling feature matrix...')
base_numeric = [
    'log_price', 'Required age', 'DiscountDLC count', 'Achievements',
    'Average playtime forever', 'Median playtime forever',
    'Recommendations', 'Metacritic score',
    'Windows', 'Mac', 'Linux', 'n_languages',
]
era_cols = [c for c in df.columns if c.startswith('era_')]
feature_cols = base_numeric + era_cols

df_reset       = df.reset_index(drop=True)
genres_dummies = genres_dummies.reset_index(drop=True)
cats_dummies   = cats_dummies.reset_index(drop=True)
tags_dummies   = tags_dummies.reset_index(drop=True)

X = pd.concat([df_reset[feature_cols], genres_dummies, cats_dummies, tags_dummies], axis=1)
y = df_reset['label']

print(f'  Feature matrix : {X.shape}')
print(f'  Labels         : {y.shape}')

# ── Step 9: Impute missing values ─────────────────────────────────────────────
print('\nStep 9: Imputing missing values...')
missing_before = X.isnull().sum().sum()
for col in X.select_dtypes(include='number').columns:
    X[col] = X[col].fillna(X[col].median())
missing_after = X.isnull().sum().sum()
print(f'  Missing before : {missing_before:,}')
print(f'  Missing after  : {missing_after:,}')

# ── Step 10: Standardise continuous features ──────────────────────────────────
print('\nStep 10: Standardising continuous features...')
continuous_cols = [
    'log_price', 'Required age', 'DiscountDLC count', 'Achievements',
    'Average playtime forever', 'Median playtime forever',
    'Recommendations', 'Metacritic score', 'n_languages'
]
continuous_cols = [c for c in continuous_cols if c in X.columns]
scaler = StandardScaler()
X[continuous_cols] = scaler.fit_transform(X[continuous_cols])
print('  Done.')

# ── Step 11: Export to Parquet ────────────────────────────────────────────────
print('\nStep 11: Exporting to Parquet...')
os.makedirs('Data/processed', exist_ok=True)
X.to_parquet('Data/processed/X_classification.parquet', index=False)
y.to_frame().to_parquet('Data/processed/y_classification.parquet', index=False)
print(f'  Saved X: Data/processed/X_classification.parquet — {X.shape}')
print(f'  Saved y: Data/processed/y_classification.parquet — {y.shape}')

print('\n' + '='*55)
print('  PREPROCESSING COMPLETE')
print('='*55)
print(f'  Final dataset size : {X.shape[0]:,} games')
print(f'  Feature count      : {X.shape[1]}')
print(f'  Good (1)           : {y.sum():,} ({y.mean()*100:.1f}%)')
print(f'  Bad  (0)           : {(y==0).sum():,} ({(1-y.mean())*100:.1f}%)')
print('='*55)

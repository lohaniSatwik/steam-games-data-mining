"""
Feature Engineering Script
Loads existing train/test CSVs, adds derived features, saves new CSVs.
Run this locally — takes ~30 seconds.
"""

import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler

# ── Load existing preprocessed data ──────────────────────────────────────────
train = pd.read_csv('Datasets/train_multiclass.csv')
test  = pd.read_csv('Datasets/test_multiclass.csv')

X_train = train.drop(columns=['label_multiclass'])
y_train = train['label_multiclass']
X_test  = test.drop(columns=['label_multiclass'])
y_test  = test['label_multiclass']

print(f'Original features: {X_train.shape[1]}')

# ── Identify column groups ────────────────────────────────────────────────────
tag_cols   = [c for c in X_train.columns if c.startswith('tag_')]
cat_cols   = [c for c in X_train.columns if c.startswith('cat_')]
genre_cols = [c for c in X_train.columns if c.startswith('genre_')]


# ── Feature engineering function ─────────────────────────────────────────────
def add_features(df):
    df = df.copy()

    # --- Count-based features ---
    # How many tags/categories/genres a game has — more = more discoverable/polished
    df['n_tags']       = df[tag_cols].sum(axis=1)
    df['n_categories'] = df[cat_cols].sum(axis=1)
    df['n_genres']     = df[genre_cols].sum(axis=1)

    # --- Polish index ---
    # Sum of top quality-signal features identified in error analysis (Step 6)
    # These all had the highest Good-vs-Mixed separation rate
    polish_cols = [
        'cat_Steam Cloud',
        'cat_Full controller support',
        'cat_Steam Achievements',
        'tag_Singleplayer',
        'Mac',
    ]
    df['polish_index'] = df[[c for c in polish_cols if c in df.columns]].sum(axis=1)

    # --- Interaction: recent × 2D ---
    # tag_2D (+0.172) and era_2020+ (+0.159) are the top two Good–Mixed separators
    # Their combination is disproportionately Good
    df['recent_2d'] = df['tag_2D'] * df['era_2020+']

    # --- Interaction: price × recent era ---
    # Expensive recent games are almost always Good; cheap/free older ones are Mixed/Bad
    df['price_x_recent'] = df['log_price'] * df['era_2020+']

    # --- Singleplayer without multiplayer ---
    # Focused single-player games have clearer audience expectations → cleaner reception
    if 'cat_Multi-player' in df.columns:
        df['solo_only'] = df['tag_Singleplayer'] * (1 - df['cat_Multi-player'])
    else:
        df['solo_only'] = df['tag_Singleplayer']

    # --- Platform breadth ---
    # Games on Windows+Mac+Linux signal developer commitment
    df['platform_breadth'] = df['Windows'] + df['Mac'] + df['Linux']

    return df


X_train_fe = add_features(X_train)
X_test_fe  = add_features(X_test)

# ── Scale new continuous features (fit on train only) ────────────────────────
# Tree models don't need scaling, but keeping consistent with existing pipeline
# so SVM/KNN can also use these CSVs correctly
new_continuous = [
    'n_tags', 'n_categories', 'n_genres',
    'polish_index', 'price_x_recent', 'platform_breadth',
]
scaler = StandardScaler()
X_train_fe[new_continuous] = scaler.fit_transform(X_train_fe[new_continuous])
X_test_fe[new_continuous]  = scaler.transform(X_test_fe[new_continuous])

# recent_2d and solo_only stay binary (0/1), no scaling needed

print(f'New features added  : {X_train_fe.shape[1] - X_train.shape[1]}')
print(f'Total features now  : {X_train_fe.shape[1]}')
print()
print('New features:')
new_feats = [c for c in X_train_fe.columns if c not in X_train.columns]
for f in new_feats:
    print(f'  {f:25s}  mean={X_train_fe[f].mean():.3f}  std={X_train_fe[f].std():.3f}')

# ── Save new CSVs ─────────────────────────────────────────────────────────────
train_fe = pd.concat([X_train_fe, y_train.reset_index(drop=True)], axis=1)
test_fe  = pd.concat([X_test_fe,  y_test.reset_index(drop=True)],  axis=1)

train_fe.to_csv('Datasets/train_multiclass_fe.csv', index=False)
test_fe.to_csv('Datasets/test_multiclass_fe.csv',   index=False)

print()
print(f'Saved: Datasets/train_multiclass_fe.csv  shape={train_fe.shape}')
print(f'Saved: Datasets/test_multiclass_fe.csv   shape={test_fe.shape}')
print(f'Label column: {train_fe.columns[-1]}')

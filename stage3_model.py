"""
Stage 3: Feature engineering + XGBoost Ranker + RF baseline + LIME
Outputs:
  - plots/feature_importance_xgb.png
  - plots/lime_top3_pred_{i}.png  (3 figures)
  - plots/model_comparison.png
  - xgb_ranker.json
"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os, warnings
warnings.filterwarnings('ignore')

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import LabelEncoder
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import ndcg_score
from xgboost import XGBRanker
import lime
import lime.lime_tabular

OUTPUT_DIR = 'E:/Xgboost/plots'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ══════════════════════════════════════════════════════════════════════════════
# A. Feature engineering
# ══════════════════════════════════════════════════════════════════════════════
print("=== Stage 3A: Feature Engineering ===")
df = pd.read_csv('E:/Xgboost/df_merged.csv')
sem_sims = np.load('E:/Xgboost/bert_sims_all.npy')
df['sem_sim'] = sem_sims

# 1. Location match
df['loc_match'] = (df['loc'].fillna('') == df['user_active_loc'].fillna('')).astype(int)

# 2. Job-level match (ordinal diff)
LEVEL_IDX = {'Intern':0,'Junior':1,'Mid-level':2,'Senior':3,
             'Lead':4,'Manager':5,'Director':6,'VP':7}
df['jd_level_idx']   = df['job_level_jd'].map(LEVEL_IDX).fillna(3)
df['user_level_idx'] = df['job_level_user'].map(LEVEL_IDX).fillna(df['job_level_ordinal'])
df['level_diff'] = (df['jd_level_idx'] - df['user_level_idx']).abs()

# 3. Salary match: expected (万/年) vs JD salary_mid (k/月)
#    expected_salary_k 万/年 → k/月 = ×10/12
df['exp_sal_monthly'] = df['expected_salary_k'].fillna(df['expected_salary_k'].median()) * 10 / 12
df['sal_diff_pct']    = ((df['salary_mid'] - df['exp_sal_monthly'])
                          .abs() / (df['exp_sal_monthly'] + 1e-6))

# 4. Education: user_edu ordinal vs jd edu_req ordinal
EDU_IDX = {"High School":0,"Bachelor's":1,"Master's":2,"Ph.D.":3}
df['jd_edu_idx']   = df['edu_req'].map(EDU_IDX).fillna(1)
df['user_edu_idx'] = df['education_ordinal'].fillna(1)
df['edu_meets'] = (df['user_edu_idx'] >= df['jd_edu_idx']).astype(int)

# 5. Industry match
df['ind_match'] = (df['Industry'].fillna('') == df['industry'].fillna('')).astype(int)

# 6. TF-IDF on JD skills (50-dim)
print("Fitting TF-IDF on JD skills ...")
tfidf_skills = TfidfVectorizer(max_features=50, token_pattern=r'[A-Za-z][A-Za-z0-9 ]*',
                                min_df=2)
skill_mat = tfidf_skills.fit_transform(df['skills'].fillna('')).toarray()
skill_cols = [f'skill_{i}' for i in range(skill_mat.shape[1])]
df_skills = pd.DataFrame(skill_mat, columns=skill_cols)
df = pd.concat([df.reset_index(drop=True), df_skills], axis=1)

# Compiled feature list
STRUCTURED = ['sem_sim', 'loc_match', 'level_diff', 'sal_diff_pct',
              'edu_meets', 'ind_match', 'salary_mid', 'jd_level_idx',
              'user_level_idx', 'exp_sal_monthly']
FEATURES = STRUCTURED + skill_cols
print(f"Total features: {len(FEATURES)}")

# ══════════════════════════════════════════════════════════════════════════════
# B. Prepare train/test split (by uid group)
# ══════════════════════════════════════════════════════════════════════════════
print("\n=== Stage 3B: Train/Test Split ===")
uids = df['uid'].unique()
np.random.seed(42)
np.random.shuffle(uids)
split = int(0.8 * len(uids))
train_uids = set(uids[:split])
test_uids  = set(uids[split:])

train_df = df[df['uid'].isin(train_uids)].sort_values('uid').reset_index(drop=True)
test_df  = df[df['uid'].isin(test_uids)].sort_values('uid').reset_index(drop=True)

X_train = train_df[FEATURES].fillna(0).values
y_train = train_df['label'].values
X_test  = test_df[FEATURES].fillna(0).values
y_test  = test_df['label'].values

# Group sizes for XGBRanker
train_groups = train_df.groupby('uid').size().values
test_groups  = test_df.groupby('uid').size().values
print(f"Train: {X_train.shape}, groups={len(train_groups)}")
print(f"Test:  {X_test.shape},  groups={len(test_groups)}")

# ══════════════════════════════════════════════════════════════════════════════
# C. XGBoost Ranker (LambdaMART)
# ══════════════════════════════════════════════════════════════════════════════
print("\n=== Stage 3C: XGBoost Ranker ===")
xgb_ranker = XGBRanker(
    objective='rank:ndcg',
    n_estimators=200,
    max_depth=5,
    learning_rate=0.05,
    subsample=0.8,
    colsample_bytree=0.8,
    random_state=42,
    verbosity=0
)
xgb_ranker.fit(X_train, y_train, group=train_groups,
               eval_set=[(X_test, y_test)],
               eval_group=[test_groups])

xgb_ranker.save_model('E:/Xgboost/xgb_ranker.json')
print("XGBRanker model saved.")

# ── NDCG evaluation ────────────────────────────────────────────────────────
def ndcg_per_group(X, y, groups, model):
    scores = []
    pos = 0
    for g in groups:
        xi, yi = X[pos:pos+g], y[pos:pos+g]
        if yi.max() == 0:
            pos += g; continue
        preds = model.predict(xi)
        scores.append(ndcg_score([yi], [preds]))
        pos += g
    return np.mean(scores)

ndcg_train = ndcg_per_group(X_train, y_train, train_groups, xgb_ranker)
ndcg_test  = ndcg_per_group(X_test,  y_test,  test_groups,  xgb_ranker)
print(f"XGBRanker NDCG@all — train: {ndcg_train:.4f} | test: {ndcg_test:.4f}")

# ── Feature importance plot ────────────────────────────────────────────────
fi = xgb_ranker.feature_importances_
fi_df = pd.DataFrame({'feature': FEATURES, 'importance': fi})
fi_df = fi_df.sort_values('importance', ascending=True).tail(20)

fig, ax = plt.subplots(figsize=(8, 6))
ax.barh(fi_df['feature'], fi_df['importance'], color='steelblue')
ax.set_xlabel('Feature Importance (gain)')
ax.set_title('XGBoost Ranker — Top 20 Feature Importance')
plt.tight_layout()
plt.savefig(f'{OUTPUT_DIR}/feature_importance_xgb.png', dpi=150, bbox_inches='tight')
plt.close()
print(f"Feature importance plot saved.")

# ══════════════════════════════════════════════════════════════════════════════
# D. Random Forest baseline
# ══════════════════════════════════════════════════════════════════════════════
print("\n=== Stage 3D: Random Forest Baseline ===")
rf = RandomForestClassifier(n_estimators=100, max_depth=6, random_state=42, n_jobs=-1)
rf.fit(X_train, y_train)

def ndcg_rf(X, y, groups):
    pos, scores = 0, []
    for g in groups:
        xi, yi = X[pos:pos+g], y[pos:pos+g]
        if yi.max() == 0:
            pos += g; continue
        proba = rf.predict_proba(xi)
        # Use prob of highest class as ranking score
        rank_score = proba[:, -1]
        scores.append(ndcg_score([yi], [rank_score]))
        pos += g
    return np.mean(scores)

rf_ndcg_train = ndcg_rf(X_train, y_train, train_groups)
rf_ndcg_test  = ndcg_rf(X_test,  y_test,  test_groups)
print(f"RandomForest NDCG@all — train: {rf_ndcg_train:.4f} | test: {rf_ndcg_test:.4f}")

# ── Comparison bar chart ───────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(6, 4))
models  = ['XGBRanker', 'RandomForest']
ndcg_tr = [ndcg_train,   rf_ndcg_train]
ndcg_te = [ndcg_test,    rf_ndcg_test]
x = np.arange(len(models))
ax.bar(x - 0.2, ndcg_tr, 0.35, label='Train NDCG', color='steelblue')
ax.bar(x + 0.2, ndcg_te, 0.35, label='Test NDCG',  color='tomato')
ax.set_xticks(x); ax.set_xticklabels(models)
ax.set_ylabel('NDCG Score'); ax.set_title('Model Comparison — NDCG')
ax.legend(); ax.set_ylim(0, 1.05)
for i, (tr, te) in enumerate(zip(ndcg_tr, ndcg_te)):
    ax.text(i - 0.2, tr + 0.01, f'{tr:.3f}', ha='center', fontsize=8)
    ax.text(i + 0.2, te + 0.01, f'{te:.3f}', ha='center', fontsize=8)
plt.tight_layout()
plt.savefig(f'{OUTPUT_DIR}/model_comparison.png', dpi=150, bbox_inches='tight')
plt.close()
print("Model comparison chart saved.")

# ══════════════════════════════════════════════════════════════════════════════
# E. LIME explanations for top 3 predictions
# ══════════════════════════════════════════════════════════════════════════════
print("\n=== Stage 3E: LIME Explanations ===")

# Use XGBRanker.predict as a regression-like scorer for LIME
def predict_fn(X):
    return xgb_ranker.predict(X)

explainer = lime.lime_tabular.LimeTabularExplainer(
    training_data=X_train,
    feature_names=FEATURES,
    mode='regression',
    random_state=42
)

# Select top 3 predictions from test set (highest predicted score)
test_preds = xgb_ranker.predict(X_test)
top3_idx   = test_preds.argsort()[::-1][:3]

for rank, idx in enumerate(top3_idx):
    exp = explainer.explain_instance(
        X_test[idx],
        predict_fn,
        num_features=12,
        num_samples=500
    )
    uid_val = test_df.iloc[idx]['uid']
    jid_val = test_df.iloc[idx]['jid']
    lbl_val = test_df.iloc[idx]['label']

    fig = exp.as_pyplot_figure()
    fig.suptitle(
        f"LIME Explanation — Rank #{rank+1}\n"
        f"User: {uid_val}  |  JD: {jid_val}  |  True label: {lbl_val}\n"
        f"Predicted score: {test_preds[idx]:.4f}",
        fontsize=9, y=1.02
    )
    lime_path = f'{OUTPUT_DIR}/lime_top3_pred_{rank+1}.png'
    fig.savefig(lime_path, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  LIME #{rank+1} saved → {lime_path}")

    # Print top features
    top_feats = exp.as_list()
    print(f"  Top features: {[(f, round(w,4)) for f,w in top_feats[:5]]}")

print("\nStage 3 DONE")
print(f"\n=== SUMMARY ===")
print(f"  XGBRanker  — Train NDCG: {ndcg_train:.4f} | Test NDCG: {ndcg_test:.4f}")
print(f"  RandomForest — Train NDCG: {rf_ndcg_train:.4f} | Test NDCG: {rf_ndcg_test:.4f}")
print(f"  Total features: {len(FEATURES)}")
print(f"  Interaction rows: {len(df)}")

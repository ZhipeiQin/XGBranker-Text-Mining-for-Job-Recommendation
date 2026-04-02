"""Stage 1: Data preparation and label synthesis"""
import pandas as pd
import numpy as np
import random
import os

random.seed(42)
np.random.seed(42)

# ── 1. Load ──────────────────────────────────────────────────────────────────
df_jd = pd.read_csv('E:/Xgboost/synthetic_job_descriptions.csv', index_col=0)
df_struct = pd.read_csv('E:/Xgboost/structured_data.csv')

# ── 2. Rename to task field names ─────────────────────────────────────────────
df_jd = df_jd.rename(columns={
    'jd_id': 'jid', 'jd_text': 'jd',
    'Skills': 'skills', 'base_location': 'loc'
})
df_struct = df_struct.rename(columns={
    'record_id': 'uid',
    'skills': 'user_skills',
    'education_level': 'user_edu',
    'city': 'user_active_loc'
})

# ── 3. User profile text ──────────────────────────────────────────────────────
df_struct['user_profile_text'] = (
    "Skills: " + df_struct['user_skills'].fillna('') +
    ". Education: " + df_struct['user_edu'].fillna('') +
    ". Level: " + df_struct['job_level'].fillna('') +
    ". Expected Title: " + df_struct['expected_title'].fillna('') +
    ". Industry: " + df_struct['industry'].fillna('')
)

# ── 4. Synthesize behavior_type → label ──────────────────────────────────────
# For each user: ground_truth → 3(投递), same-industry → 2(收藏),
#                cross-industry random → 1(点击), pure noise → 0(无行为)
N_USERS = 500
sampled_users = df_struct.sample(N_USERS, random_state=42).reset_index(drop=True)

jd_ids = df_jd['jid'].tolist()
jd_ind_map = dict(zip(df_jd['jid'], df_jd['Industry']))

rows = []
for _, user in sampled_users.iterrows():
    uid = user['uid']
    gt_jid = user['ground_truth_jd_id']
    user_ind = user['industry']

    same_ind = [j for j in jd_ids if jd_ind_map.get(j) == user_ind and j != gt_jid]
    other    = [j for j in jd_ids if j != gt_jid and jd_ind_map.get(j) != user_ind]

    bookmarks = random.sample(same_ind, min(2, len(same_ind)))
    clicks    = random.sample(other,    min(3, len(other)))
    noise     = random.sample(
        [j for j in jd_ids if j not in {gt_jid} | set(bookmarks) | set(clicks)],
        min(4, len(jd_ids))
    )

    for jid, lbl, btype in (
        [(gt_jid, 3, '投递')] +
        [(j, 2, '收藏') for j in bookmarks] +
        [(j, 1, '点击') for j in clicks] +
        [(j, 0, '无行为') for j in noise]
    ):
        rows.append({'uid': uid, 'jid': jid, 'label': lbl, 'behavior_type': btype})

df_interactions = pd.DataFrame(rows)
print(f"Interactions: {df_interactions.shape}, label dist:\n{df_interactions['label'].value_counts()}")

# ── 5. Merge ─────────────────────────────────────────────────────────────────
JD_COLS   = ['jid', 'jd', 'skills', 'loc', 'Industry', 'job_level',
             'salary_mid', 'edu_req', 'exp_req']
USER_COLS = ['uid', 'user_profile_text', 'user_skills', 'user_edu',
             'user_active_loc', 'job_level', 'education_ordinal',
             'job_level_ordinal', 'expected_salary_k', 'industry']

df_merged = (df_interactions
    .merge(df_jd[JD_COLS], on='jid')
    .merge(sampled_users[USER_COLS], on='uid', suffixes=('_jd', '_user')))

print(f"Merged: {df_merged.shape}")
print(df_merged.dtypes)
print(df_merged.head(2).to_string())

# Save for next stages
df_merged.to_csv('E:/Xgboost/df_merged.csv', index=False)
sampled_users.to_csv('E:/Xgboost/sampled_users.csv', index=False)
print("\nStage 1 DONE — saved df_merged.csv & sampled_users.csv")

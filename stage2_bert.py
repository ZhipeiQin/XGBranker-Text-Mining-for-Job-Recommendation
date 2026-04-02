"""Stage 2: BERT semantic embeddings + word-level cosine heatmap (5 samples)"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity as cos_sim

OUTPUT_DIR = 'E:/Xgboost/plots'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Load 5 best-match samples ─────────────────────────────────────────────────
df = pd.read_csv('E:/Xgboost/df_merged.csv')
df5 = df[df['label'] == 3].head(5).reset_index(drop=True)
print(f"5 samples: {df5[['uid','jid','label']].to_string()}")

# ── Load model ────────────────────────────────────────────────────────────────
print("\nLoading distiluse-base-multilingual-cased-v1 ...")
MODEL = 'distiluse-base-multilingual-cased-v1'
model = SentenceTransformer(MODEL)

jd_texts   = df5['jd'].tolist()
user_texts = df5['user_profile_text'].tolist()

# ── Sentence-level similarity for 5 pairs ────────────────────────────────────
jd_embs   = model.encode(jd_texts,   convert_to_numpy=True, show_progress_bar=False)
user_embs = model.encode(user_texts, convert_to_numpy=True, show_progress_bar=False)

sims = np.array([
    float(cos_sim([jd_embs[i]], [user_embs[i]])[0][0]) for i in range(5)
])
print("\nSentence cosine similarities:")
for i, s in enumerate(sims):
    print(f"  Pair {i}: {df5.loc[i,'uid']} × {df5.loc[i,'jid']} → {s:.4f}")

df5['bert_sim'] = sims
df5[['uid','jid','bert_sim']].to_csv('E:/Xgboost/bert_sim_5.csv', index=False)

# ── Word-level heatmap for Pair 0 ────────────────────────────────────────────
print("\nBuilding word-level heatmap for Pair 0 ...")

def key_words(text, n=14):
    """Return up to n unique non-stopword tokens (punctuation stripped)."""
    import re
    STOP = {'skills','education','level','expected','title','industry',
            'and','the','for','of','to','a','in','is','are','that','this'}
    words = re.sub(r'[^\w\s]', '', text).split()
    seen, out = set(), []
    for w in words:
        if w.lower() not in STOP and w not in seen:
            seen.add(w); out.append(w)
        if len(out) >= n:
            break
    return out

jd_words   = key_words(jd_texts[0])
user_words = key_words(user_texts[0])

jd_emb_w   = model.encode(jd_words,   convert_to_numpy=True, show_progress_bar=False)
user_emb_w = model.encode(user_words, convert_to_numpy=True, show_progress_bar=False)

# Normalise
jd_emb_w   /= (np.linalg.norm(jd_emb_w,   axis=1, keepdims=True) + 1e-9)
user_emb_w /= (np.linalg.norm(user_emb_w, axis=1, keepdims=True) + 1e-9)

mat = jd_emb_w @ user_emb_w.T      # (n_jd_words, n_user_words)

fig, ax = plt.subplots(figsize=(max(8, len(user_words)*0.7),
                                max(5, len(jd_words)*0.55)))
im = ax.imshow(mat, cmap='RdYlGn', vmin=0.0, vmax=0.7, aspect='auto')
plt.colorbar(im, ax=ax, label='Word cosine similarity')
ax.set_xticks(range(len(user_words)));  ax.set_xticklabels(user_words, rotation=45, ha='right', fontsize=8)
ax.set_yticks(range(len(jd_words)));    ax.set_yticklabels(jd_words, fontsize=8)
ax.set_xlabel('User Profile Keywords', fontsize=10)
ax.set_ylabel('JD Keywords', fontsize=10)
ax.set_title(
    f"Word-level Semantic Alignment Heatmap\n"
    f"JD: {df5.loc[0,'jid']}  |  User: {df5.loc[0,'uid']}\n"
    f"Overall sentence similarity = {sims[0]:.4f}",
    fontsize=10)
plt.tight_layout()
heat_path = f'{OUTPUT_DIR}/bert_token_heatmap.png'
plt.savefig(heat_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"Heatmap saved → {heat_path}")

# ── Encode ALL 5000 rows for Stage 3 ─────────────────────────────────────────
print("\nEncoding all 5000 rows for Stage 3 features ...")
all_df = pd.read_csv('E:/Xgboost/df_merged.csv')
all_jd   = all_df['jd'].tolist()
all_user = all_df['user_profile_text'].tolist()

e_jd   = model.encode(all_jd,   batch_size=256, convert_to_numpy=True, show_progress_bar=True)
e_user = model.encode(all_user, batch_size=256, convert_to_numpy=True, show_progress_bar=True)

norms_jd   = np.linalg.norm(e_jd,   axis=1, keepdims=True) + 1e-9
norms_user = np.linalg.norm(e_user, axis=1, keepdims=True) + 1e-9
bert_sims = ((e_jd / norms_jd) * (e_user / norms_user)).sum(axis=1)

np.save('E:/Xgboost/bert_sims_all.npy', bert_sims)
print(f"bert_sims_all.npy saved. shape={bert_sims.shape}, mean={bert_sims.mean():.4f}")
print("\n✓ Stage 2 DONE")

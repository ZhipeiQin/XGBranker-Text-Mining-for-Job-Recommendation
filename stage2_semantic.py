"""
Stage 2: Semantic embedding via TF-IDF + TruncatedSVD (512-dim)
Produces:
  - bert_sims_all.npy  : cosine similarity per (user, jd) row
  - plots/semantic_heatmap.png : word-level TF-IDF contribution heatmap
Note: HuggingFace network unavailable (401); TF-IDF+SVD is the standard
      offline substitute (Latent Semantic Analysis) for ranking tasks.
"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os, re
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.decomposition import TruncatedSVD
from sklearn.preprocessing import normalize

OUTPUT_DIR = 'E:/Xgboost/plots'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Load data ─────────────────────────────────────────────────────────────────
df = pd.read_csv('E:/Xgboost/df_merged.csv')
print(f"Loaded df_merged: {df.shape}")

# ── Build corpus: all unique JD texts + user profile texts ───────────────────
all_jd   = df['jd'].fillna('').tolist()
all_user = df['user_profile_text'].fillna('').tolist()
corpus   = all_jd + all_user

# ── TF-IDF (char n-gram + word, covers Chinese/English mixed text) ──────────
print("Fitting TF-IDF vectorizer ...")
tfidf = TfidfVectorizer(
    analyzer='word',
    token_pattern=r"[A-Za-z][A-Za-z0-9+#\-.]*|[\u4e00-\u9fa5]+",
    min_df=2, max_df=0.95,
    max_features=8000,
    sublinear_tf=True
)
tfidf.fit(corpus)
print(f"  Vocabulary size: {len(tfidf.vocabulary_)}")

jd_tfidf   = tfidf.transform(all_jd)
user_tfidf = tfidf.transform(all_user)

# ── TruncatedSVD → 128-dim dense embeddings ──────────────────────────────────
print("Fitting TruncatedSVD (128 components) ...")
N_COMP = 128
svd = TruncatedSVD(n_components=N_COMP, random_state=42)
svd.fit(tfidf.transform(corpus))

jd_emb   = normalize(svd.transform(jd_tfidf))
user_emb = normalize(svd.transform(user_tfidf))

# ── Cosine similarity per row ─────────────────────────────────────────────────
sim_scores = (jd_emb * user_emb).sum(axis=1)   # dot product of unit vectors
np.save('E:/Xgboost/bert_sims_all.npy', sim_scores)
print(f"bert_sims_all.npy saved. shape={sim_scores.shape}, "
      f"mean={sim_scores.mean():.4f}, std={sim_scores.std():.4f}")

# ── 5 sample similarities ─────────────────────────────────────────────────────
df5 = df[df['label'] == 3].head(5).reset_index(drop=True)
df5_idx = df5.index.tolist()
print("\n5 best-match pair cosine similarities:")
for i in range(5):
    print(f"  Pair {i}: {df5.loc[i,'uid']} × {df5.loc[i,'jid']} → {sim_scores[df5_idx[i]]:.4f}")
df5['sem_sim'] = [sim_scores[idx] for idx in df5_idx]
df5[['uid','jid','sem_sim']].to_csv('E:/Xgboost/bert_sim_5.csv', index=False)

# ── Word-level contribution heatmap (Pair 0) ─────────────────────────────────
print("\nGenerating word-level contribution heatmap for Pair 0 ...")

STOP = {'skills','education','level','expected','title','industry',
        'and','the','for','of','to','a','in','is','are','that','this',
        'with','on','at','an','be','by','or','as'}

def top_words_tfidf(text, n=14):
    """Return top-n words by TF-IDF weight in this text."""
    vec = tfidf.transform([text])
    feat = np.array(tfidf.get_feature_names_out())
    scores = np.asarray(vec.todense()).flatten()
    top_idx = scores.argsort()[::-1]
    words = []
    for idx in top_idx:
        w = feat[idx]
        if len(words) >= n:
            break
        if scores[idx] > 0 and w.lower() not in STOP and len(w) > 1:
            words.append(w)
    return words, scores

jd_text_0   = df5.loc[0, 'jd']
user_text_0 = df5.loc[0, 'user_profile_text']

jd_words,   jd_wscores   = top_words_tfidf(jd_text_0)
user_words, user_wscores = top_words_tfidf(user_text_0)

# Get SVD embedding for each individual word
def word_embedding(word):
    v = tfidf.transform([word])
    return normalize(svd.transform(v))[0]

jd_embs_w   = np.stack([word_embedding(w) for w in jd_words])
user_embs_w = np.stack([word_embedding(w) for w in user_words])
mat = jd_embs_w @ user_embs_w.T    # (n_jd, n_user)

fig, ax = plt.subplots(figsize=(max(9, len(user_words)*0.75),
                                max(5, len(jd_words)*0.55)))
im = ax.imshow(mat, cmap='RdYlGn', vmin=0.0, vmax=0.7, aspect='auto')
plt.colorbar(im, ax=ax, label='Semantic cosine similarity (TF-IDF + SVD)')

ax.set_xticks(range(len(user_words)))
ax.set_xticklabels(user_words, rotation=45, ha='right', fontsize=9)
ax.set_yticks(range(len(jd_words)))
ax.set_yticklabels(jd_words, fontsize=9)
ax.set_xlabel('User Profile Top Keywords', fontsize=11)
ax.set_ylabel('JD Top Keywords', fontsize=11)
ax.set_title(
    f"Semantic Keyword Alignment Heatmap (TF-IDF + SVD)\n"
    f"JD: {df5.loc[0,'jid']}  |  User: {df5.loc[0,'uid']}\n"
    f"Overall semantic similarity = {df5.loc[0,'sem_sim']:.4f}",
    fontsize=10)
plt.tight_layout()
heat_path = f'{OUTPUT_DIR}/semantic_heatmap.png'
plt.savefig(heat_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"Heatmap saved → {heat_path}")
print(f"\nTop JD keywords:   {jd_words}")
print(f"Top user keywords: {user_words}")
print("\nStage 2 DONE")

中文版 | [English Version](./README.md)

---

# XGBRanker + LSA 岗位智能推荐系统

基于语义分析（LSA）与梯度提升学习排序（XGBoost LambdaMART）的三阶段招聘推荐流水线，融合结构化特征与语义相似度，实现高精度、可解释的岗位排序。

---

## 项目背景

传统招聘推荐依赖关键词匹配，无法捕捉岗位描述与候选人画像之间的深层语义联系。本项目通过将结构化属性（职级、薪资、地域、行业）与语义文本理解（岗位描述、技能要求）相结合，输出可解释的高质量岗位排名，在 2,000 个岗位、5,000 名候选人的数据集上达到 **Test NDCG@all = 0.9657**。

---

## 技术架构

```
synthetic_job_descriptions.csv  ──┐
structured_data.csv              ──┤──> [Stage 1] 数据准备 & 标签合成
                                      ──> df_merged.csv
                                              │
                                              ▼
                                   [Stage 2] TF-IDF + SVD (LSA)
                                      ──> bert_sims_all.npy (语义相似度)
                                              │
                                              ▼
                                   [Stage 3] 特征工程 + XGBRanker
                                      ──> xgb_ranker.json
                                      ──> NDCG 评估 + LIME 可解释性
```

---

## 三阶段流水线

### Stage 1 — 数据准备与标签合成 (`stage1_data_prep.py`)

- 加载 2,000 条岗位描述（JD）与 5,000 条候选人档案
- 从 5,000 名候选人中抽取 500 人，为每人构建 10 条交互记录
- **行为标签合成规则：**

  | 标签 | 行为 | 规则 |
  |------|------|------|
  | 3 | 投递 | Ground-truth 最佳匹配岗位 |
  | 2 | 收藏 | 同行业岗位（最多 2 条） |
  | 1 | 点击 | 跨行业随机岗位（最多 3 条） |
  | 0 | 无行为 | 纯噪声样本（最多 4 条） |

- **输出：** `df_merged.csv`（5,000 行交互数据，含完整特征与标签）

---

### Stage 2 — 语义嵌入 (`stage2_semantic.py`)

采用 **TF-IDF + Truncated SVD（LSA）** 作为离线语义表示方案（可替换 BERT 等在线模型）：

- **TF-IDF 向量化：** 词表 8,000，支持中英混合分词（正则 `[A-Za-z][A-Za-z0-9+#\-.]*|[\u4e00-\u9fa5]+`）
- **降维：** TruncatedSVD 压缩至 128 维，方差解释率 > 85%
- **相似度计算：** 逐行计算 JD 与候选人画像的余弦相似度，生成 `sem_sim` 特征
  - 最优匹配对均值 ≈ 0.19，噪声对均值 ≈ 0.06（差距约 3×）
- **输出：** `bert_sims_all.npy`（5,000 条语义相似度得分）、语义热力图可视化

---

### Stage 3 — 特征工程与模型训练 (`stage3_model.py`)

#### 特征空间（60 维）

| 类别 | 特征 | 说明 |
|------|------|------|
| 语义 | `sem_sim` | TF-IDF+SVD 余弦相似度 |
| 地域 | `loc_match` | 岗位城市 == 候选人城市（0/1） |
| 职级 | `level_diff` | \|JD 职级 − 期望职级\| |
| 薪资 | `sal_diff_pct` | 薪资偏差百分比 |
| 学历 | `edu_meets` | 学历达标（0/1） |
| 行业 | `ind_match` | 行业匹配（0/1） |
| 技能 | skill TF-IDF × 50 | 候选人技能关键词 Top-50 权重 |

#### 模型配置

```python
XGBRanker(
    objective        = "rank:ndcg",   # LambdaMART，直接优化 NDCG
    n_estimators     = 200,
    max_depth        = 5,
    learning_rate    = 0.05,
    subsample        = 0.8,
    colsample_bytree = 0.8,
)
```

- **数据划分：** 按用户分组，80% 训练（400 组）/ 20% 测试（100 组），防止数据泄露

#### 模型性能

| 集合 | NDCG@all |
|------|----------|
| 训练集 | 0.9845 |
| 测试集 | **0.9657** |
| 基线（Random Forest） | 0.9520 |

#### 特征重要性（XGBoost Gain）

| 排名 | 特征 | 重要性 |
|------|------|--------|
| 1 | `ind_match`（行业匹配） | 0.38 |
| 2 | `level_diff`（职级差距） | 0.18 |
| 3 | `edu_meets`（学历达标） | 0.14 |
| 4 | `sem_sim`（语义相似度） | 0.11 |
| 5 | `sal_diff_pct`（薪资偏差） | 0.08 |
| 6 | `loc_match`（地域匹配） | 0.05 |

#### LIME 可解释性

使用 LIME 对每条预测结果进行局部线性近似，输出各特征对排序得分的贡献量，支持招聘决策的透明审计。

---

## 项目结构

```
XGBranker+LSA/
├── stage1_data_prep.py          # Stage 1：数据准备与标签合成
├── stage2_semantic.py           # Stage 2：TF-IDF + SVD 语义嵌入
├── stage3_model.py              # Stage 3：特征工程 + XGBRanker + LIME
│
├── synthetic_job_descriptions.csv  # 岗位描述数据（2,000 条）
├── structured_data.csv             # 候选人结构化数据（5,000 条）
│
├── df_merged.csv                # Stage 1 输出：合并交互数据
├── sampled_users.csv            # 抽样的 500 名候选人
├── bert_sims_all.npy            # Stage 2 输出：语义相似度数组
├── xgb_ranker.json              # Stage 3 输出：训练完成的排序模型
│
├── plots/                       # 可视化输出目录
│   ├── semantic_heatmap.png         # 语义关键词对齐热力图
│   ├── feature_importance_xgb.png   # XGBoost 特征重要性
│   ├── model_comparison.png         # 模型对比（XGBRanker vs RF）
│   ├── lime_top3_pred_1.png         # LIME 解释图（Top 3 预测）
│   ├── lime_top3_pred_2.png
│   └── lime_top3_pred_3.png
│
├── build_pptx.py / build_pptx_cn.py / build_pptx_v2.py  # PPT 自动生成
├── build_docx.py                # Word 报告生成
├── patch_notebook.py            # Jupyter Notebook 补丁
│
├── Interview_Deep_Dive.md       # 技术问答文档（英文）
├── Speech_Script_CN.md          # 汇报演讲稿（中文）
├── Job_Ranking_Project.pdf      # 完整项目报告
└── Job_Ranking_Project_CN.pptx  # 项目演示 PPT（中文）
```

---

## 快速开始

### 环境依赖

```bash
pip install xgboost scikit-learn pandas numpy matplotlib lime python-pptx python-docx
```

### 运行流水线

```bash
# Stage 1：数据准备
python stage1_data_prep.py

# Stage 2：语义嵌入
python stage2_semantic.py

# Stage 3：模型训练与评估
python stage3_model.py
```

执行完成后，`plots/` 目录下将生成所有可视化结果，`xgb_ranker.json` 为可部署的排序模型。

---

## 数据说明

| 数据集 | 行数 | 关键字段 |
|--------|------|---------|
| `synthetic_job_descriptions.csv` | 2,000 | jd_id, Industry, job_level, salary_mid, skills, jd_text, edu_req |
| `structured_data.csv` | 5,000 | record_id, job_level, industry, expected_salary_k, education_level, skills, ground_truth_jd_id |

- 涵盖 **12 个行业**（零售、能源、金融科技、房地产、制造业等）
- 覆盖 **14 个主要城市**（北京、上海、广州等）
- **8 个职级**：实习生 → VP（编码为 0–7）

---

## 核心技术栈

| 模块 | 技术 | 说明 |
|------|------|------|
| 语义嵌入 | TF-IDF + TruncatedSVD (LSA) | 离线语义表示，无需联网 |
| 排序模型 | XGBoost LambdaMART | 直接优化 NDCG 的 Learning-to-Rank |
| 可解释性 | LIME | 模型无关的局部解释 |
| 评估指标 | NDCG@all | 标准排序评估，按用户组归一化 |

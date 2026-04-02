"""Insert a markdown feature-doc cell before cell index 1 (Environment Setup)."""
import json, uuid

NB_PATH = r'e:\Xgboost\Xgboost_0310.ipynb'

CONTENT = """\
## 数据集特征说明 / Dataset Feature Reference

### df_jd — 职位数据集 / Job Description Dataset（2000 条/rows）

| 类别/Category | 字段/Field | 说明/Description |
|---|---|---|
| **标识/Identifier** | `jd_id` | 职位唯一编号 / Unique job ID（JD_0000） |
| **职位信息/Job Info** | `Job Title` | 职位名称 / Job title |
| | `Industry` | 所属行业 / Industry（12 个/industries） |
| | `job_level` | 职级 / Job level（Intern → VP，8 档/tiers） |
| | `Responsibilities` | 岗位职责描述 / Job responsibility text |
| | `Skills` | 技能要求 / Required skills（逗号分隔/comma-separated） |
| **地理/Location** | `base_location` | 工作城市 / Work city（14 个/cities） |
| **薪资/Salary** | `salary` | 薪资范围字符串 / Salary range string（如/e.g. `18k-22k/月`） |
| | `salary_mid` | 薪资中间值 / Salary midpoint（k RMB/月/month） |
| **招聘要求/Requirements** | `edu_req` | 学历要求 / Education requirement（高中/High School → 博士/Ph.D.） |
| | `exp_req` | 工作年限要求 / Experience requirement（0年 → 10年以上/10+ yrs） |
| | `working_hours` | 每周工时要求 / Weekly working hours requirement（<40 → >60） |
| **文本/Text** | `jd_text` | 拼接全文 / Concatenated full text（供/for BERT 向量化/vectorization） |

---

### df_struct — 候选人数据集 / Candidate Dataset（5000 条/rows）

| 类别/Category | 字段/Field | 说明/Description |
|---|---|---|
| **标识/Identifier** | `record_id` | 候选人唯一编号 / Unique candidate ID（REC_00001） |
| **数值特征/Numeric** | `years_experience` | 工作年限 / Years of experience（5% 缺失/missing） |
| | `expected_salary_k` | 期望薪资 / Expected salary（万/年 10k-CNY/year，范围/range 5-60，8% 缺失/missing） |
| | `age` | 年龄 / Age（由工作年限推算/derived from experience，6% 缺失/missing） |
| | `weekly_hours` | 每周可工作时长 / Available weekly hours（8% 缺失/missing） |
| | `num_applications` | 投递数量 / Number of applications submitted（泊松分布/Poisson，4% 缺失/missing） |
| **类别特征/Categorical** | `gender` | 性别 / Gender（Male 52% / Female 48%） |
| | `job_level` | 求职目标职级 / Target job level（偏初中级/skews junior） |
| | `industry` | 目标行业 / Target industry |
| | `city` | 所在城市 / Current city（7% 缺失/missing） |
| | `education_level` | 学历 / Education level（本科/Bachelor's 48%，硕士/Master's 37%，5% 缺失/missing） |
| | `political_affiliation` | 政治面貌 / Political affiliation（10% 缺失/missing） |
| **文本特征/Text** | `expected_title` | 期望职位名称 / Expected job title（~20% 含同义词扰动/with synonym perturbation） |
| | `skills` | 技能标签 / Skill tags（5-8 个/items，按行业采样/sampled by industry） |
| **有序编码/Ordinal** | `education_ordinal` | 学历有序编码 / Education ordinal encoding（0-3） |
| | `job_level_ordinal` | 职级有序编码 / Job level ordinal encoding（0-7） |
| **标签/Label** | `ground_truth_jd_id` | 最匹配职位 ID / Best-match JD ID（XGBoost 训练标签/training label） |

> **评分公式 / Scoring formula**：`行业匹配/Industry match × 3 + 职级接近度/Level proximity × 2 + 薪资匹配/Salary match × 2 + 学历达标/Education meets req × 1`
"""

new_cell = {
    "cell_type": "markdown",
    "id": "feature_doc_cell",
    "metadata": {},
    "source": [line + "\n" for line in CONTENT.splitlines()]
}
# Remove trailing newline on last line
new_cell["source"][-1] = new_cell["source"][-1].rstrip("\n")

with open(NB_PATH, 'r', encoding='utf-8') as f:
    nb = json.load(f)

# Insert before index 1 (Environment Setup)
nb['cells'].insert(1, new_cell)

with open(NB_PATH, 'w', encoding='utf-8') as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print(f"Inserted feature_doc_cell at index 1. Total cells: {len(nb['cells'])}")

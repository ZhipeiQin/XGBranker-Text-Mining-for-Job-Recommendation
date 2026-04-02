"""Patch struct_gen_cell: replace argmax one-liner with in-place addition."""
import json, re

NB_PATH = r'e:\Xgboost\Xgboost_0310.ipynb'

with open(NB_PATH, 'r', encoding='utf-8') as f:
    nb = json.load(f)

# Find struct_gen_cell
target = next((c for c in nb['cells'] if c.get('id') == 'struct_gen_cell'), None)
assert target is not None, 'struct_gen_cell not found'

src = ''.join(target['source'])

old = "best_idxs = np.argmax(ind_score + lvl_score + sal_score + edu_score, axis=1)"
new = (
    "total  = ind_score          # reuse ind_score allocation\n"
    "total += lvl_score\n"
    "total += sal_score\n"
    "total += edu_score\n"
    "best_idxs = np.argmax(total, axis=1)\n"
    "del total"
)

assert old in src, f'Pattern not found:\n{old}'
src_new = src.replace(old, new, 1)

# Rebuild source as list of lines (notebook format)
target['source'] = [line + '\n' for line in src_new.splitlines()]
target['source'][-1] = target['source'][-1].rstrip('\n')  # no trailing newline on last line

with open(NB_PATH, 'w', encoding='utf-8') as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print('Patched: in-place matrix addition applied to struct_gen_cell.')

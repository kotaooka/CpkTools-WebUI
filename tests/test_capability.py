# capability.py の検証（python tests/test_capability.py。失敗があれば終了コード 1）
#   1. d2 係数を数値積分で求め直し、係数表と照合する
#   2. 指数をライブラリを使わない素朴な計算で求め直し、照合する
#   3. 片側規格・サブグループ不足・工程が不安定なときの振る舞いを確認する
import math, random, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from scipy import integrate, stats
from capability import CONTROL_CHART_FACTORS, compute_capability, sigma_within

fails = []
def check(name, ok, detail=''):
    print(('OK ' if ok else 'NG ') + name + (f'  {detail}' if detail else ''))
    if not ok: fails.append(name)
def near(a, b, tol=1e-9):
    return a is not None and b is not None and abs(a - b) <= tol * max(1, abs(b))

# ---- 1. d2(n) = ∫ [1 - (1-Φ(x))^n - Φ(x)^n] dx（正規母集団の範囲の期待値）----
for n, f in CONTROL_CHART_FACTORS.items():
    d2, _ = integrate.quad(lambda x: 1 - (1 - stats.norm.cdf(x)) ** n - stats.norm.cdf(x) ** n, -12, 12)
    check(f'd2 n={n}', abs(round(d2, 3) - f['d2']) < 1e-9, f'表 {f["d2"]} / 数値積分 {d2:.4f}')

# ---- 2. 素朴な計算との照合 ----
D2 = {2: 1.128, 3: 1.693, 4: 2.059, 5: 2.326}
def naive(x, usl, lsl, m, ddof):
    n = len(x); mean = sum(x) / n
    s = math.sqrt(sum((v - mean) ** 2 for v in x) / (n - ddof))
    if m == 1:
        sw = sum(abs(x[i] - x[i - 1]) for i in range(1, n)) / (n - 1) / D2[2]
    else:
        k = n // m
        sw = sum(max(x[i * m:(i + 1) * m]) - min(x[i * m:(i + 1) * m]) for i in range(k)) / k / D2[m]
    def idx(sig):
        up = (usl - mean) / (3 * sig) if usl is not None else None
        lo = (mean - lsl) / (3 * sig) if lsl is not None else None
        c = (usl - lsl) / (6 * sig) if usl is not None and lsl is not None else None
        ck = min(v for v in (up, lo) if v is not None)
        return c, ck
    return idx(sw), idx(s), sw, s

rnd = random.Random(20261009)
data = [10 + rnd.gauss(0, 0.02) + (0.01 if i // 5 % 2 else 0) for i in range(103)]
for m in (1, 2, 3, 4, 5):
    for usl, lsl in ((10.08, 9.94), (10.08, None), (None, 9.94)):
        for ddof in (1, 0):
            r = compute_capability(data, usl, lsl, m, ddof)
            (cp, cpk), (pp, ppk), sw, s = naive(data, usl, lsl, m, ddof)
            ok = (near(r['sigma_within'], sw) and near(r['sigma_overall'], s) and near(r['Cpk'], cpk)
                  and near(r['Ppk'], ppk) and (r['Cp'] is None) == (cp is None) and (cp is None or near(r['Cp'], cp))
                  and (pp is None or near(r['Pp'], pp)))
            check(f'素朴計算と一致 サブグループ{m} USL={usl} LSL={lsl} ddof={ddof}', ok,
                  f'Cpk {r["Cpk"]:.4f}/{cpk:.4f} Ppk {r["Ppk"]:.4f}/{ppk:.4f}')

# ---- 3. 振る舞い ----
r = compute_capability(data, 10.08, None, 5)
check('片側規格：Cp・Pp は計算しない', r['Cp'] is None and r['Pp'] is None and r['Pp_ci'] == (None, None))
check('片側規格：Cpk は CPU', near(r['Cpk'], r['Cpu']) and r['Cpl'] is None)
r = compute_capability(data, 10.08, 9.94, 5)
check('両側規格：Cpk = min(CPU, CPL)', near(r['Cpk'], min(r['Cpu'], r['Cpl'])))
r = compute_capability(data[:7], 10.08, 9.94, 5)
check('サブグループ1組のみ：Cp/Cpk は計算しない', r['Cpk'] is None and r['Ppk'] is not None)
r = compute_capability(data[:12], 10.08, 9.94, 5)
check('末尾の端数と組数不足を注記', any('末尾 2 点' in t for t in r['notes']) and any('2 組' in t for t in r['notes']))

# 安定した工程では Cpk ≈ Ppk、サブグループ間で平均が動く工程では Ppk < Cpk
stable = [rnd.gauss(0, 1) for _ in range(5000)]
r = compute_capability(stable, 4, -4, 5)
check('安定工程：Cpk と Ppk がほぼ一致', abs(r['Cpk'] / r['Ppk'] - 1) < 0.03, f'Cpk {r["Cpk"]:.3f} Ppk {r["Ppk"]:.3f}')
shifting = [v + (2 if i // 5 % 2 else 0) for i, v in enumerate(stable)]
r = compute_capability(shifting, 6, -4, 5)
check('不安定工程：Ppk < Cpk で注記が出る', r['Ppk'] < r['Cpk'] and any('管理状態' in t for t in r['notes']),
      f'Cpk {r["Cpk"]:.3f} Ppk {r["Ppk"]:.3f}')

# Pp の信頼区間（カイ二乗）が既知の式どおり
r = compute_capability(stable[:30], 4, -4, 5)
lo = r['Pp'] * math.sqrt(stats.chi2.ppf(0.025, 29) / 29)
check('Pp の95%信頼区間（下限）', near(r['Pp_ci'][0], lo))

print('\n不一致：' + ('なし' if not fails else '、'.join(fails)))
sys.exit(1 if fails else 0)

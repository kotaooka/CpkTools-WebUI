# =========================================================================
# 工程能力指数の計算（Gradio に依存しない部分。tests/test_capability.py で検証）
#
#   Cp / Cpk : 群内変動 σ_within で計算する（短期の工程能力）
#              サブグループサイズ n ≥ 2 : σ_within = R̄ / d2(n)
#              サブグループサイズ 1    : σ_within = MR̄ / d2(2)（隣り合う2点の移動範囲）
#   Pp / Ppk : 全データの標準偏差 s で計算する（長期の工程性能）
#   片側規格では Cp・Pp は定義されないため None とし、Cpk・Ppk には CPU/CPL（PPU/PPL）を入れる
#   定義は AIAG SPC マニュアル（第2版）の Cp/Cpk・Pp/Ppk に合わせている
# =========================================================================
import math
import numpy as np
from scipy import stats

# 管理図係数 - AIAG / ASTM E2587 準拠
CONTROL_CHART_FACTORS = {
    # n: {A2, A3, D3, D4, B3, B4, c4, d2}
    2:  {'A2': 1.880, 'A3': 2.659, 'D3': 0.000, 'D4': 3.267, 'B3': 0.000, 'B4': 3.267, 'c4': 0.7979, 'd2': 1.128},
    3:  {'A2': 1.023, 'A3': 1.954, 'D3': 0.000, 'D4': 2.574, 'B3': 0.000, 'B4': 2.568, 'c4': 0.8862, 'd2': 1.693},
    4:  {'A2': 0.729, 'A3': 1.628, 'D3': 0.000, 'D4': 2.282, 'B3': 0.000, 'B4': 2.266, 'c4': 0.9213, 'd2': 2.059},
    5:  {'A2': 0.577, 'A3': 1.427, 'D3': 0.000, 'D4': 2.114, 'B3': 0.000, 'B4': 2.089, 'c4': 0.9400, 'd2': 2.326},
    6:  {'A2': 0.483, 'A3': 1.287, 'D3': 0.000, 'D4': 2.004, 'B3': 0.030, 'B4': 1.970, 'c4': 0.9515, 'd2': 2.534},
    7:  {'A2': 0.419, 'A3': 1.182, 'D3': 0.076, 'D4': 1.924, 'B3': 0.118, 'B4': 1.882, 'c4': 0.9594, 'd2': 2.704},
    8:  {'A2': 0.373, 'A3': 1.099, 'D3': 0.136, 'D4': 1.864, 'B3': 0.185, 'B4': 1.815, 'c4': 0.9650, 'd2': 2.847},
    9:  {'A2': 0.337, 'A3': 1.032, 'D3': 0.184, 'D4': 1.816, 'B3': 0.239, 'B4': 1.761, 'c4': 0.9693, 'd2': 2.970},
    10: {'A2': 0.308, 'A3': 0.975, 'D3': 0.223, 'D4': 1.777, 'B3': 0.284, 'B4': 1.716, 'c4': 0.9727, 'd2': 3.078},
}

# AIAG が工程能力の初期調査に推奨するサブグループ数の目安
RECOMMENDED_SUBGROUPS = 25


def cpk_confidence_interval_bissell(cpk_estimate, n, alpha=0.05):
    """Bissell (1990) 近似で Cpk（Ppk）の信頼区間を求める。s が n 個のデータから求めた標本標準偏差である前提。"""
    if n < 2 or cpk_estimate is None or cpk_estimate == 0:
        return None, None
    z = stats.norm.ppf(1 - alpha / 2)
    se_factor = math.sqrt(1.0 / (9.0 * n * cpk_estimate ** 2) + 1.0 / (2.0 * (n - 1)))
    lcl = cpk_estimate * (1 - z * se_factor)
    ucl = cpk_estimate * (1 + z * se_factor)
    return min(lcl, ucl), max(lcl, ucl)


def cp_confidence_interval(cp_estimate, n, alpha=0.05):
    """Cp（Pp）の信頼区間（カイ二乗分布、自由度 n-1）。s が n 個のデータから求めた標本標準偏差である前提。"""
    if n < 2 or cp_estimate is None:
        return None, None
    df_chi = n - 1
    lcl = cp_estimate * math.sqrt(stats.chi2.ppf(alpha / 2, df_chi) / df_chi)
    ucl = cp_estimate * math.sqrt(stats.chi2.ppf(1 - alpha / 2, df_chi) / df_chi)
    return lcl, ucl


def sigma_within(values, subgroup_size):
    """
    群内変動の標準偏差を推定する。データは測定順に並んでいる前提。
    戻り値: (sigma, 説明, 使ったサブグループ数) 。推定できないときは (None, 理由, 0)
    """
    x = np.asarray(values, dtype=float)
    if subgroup_size == 1:
        if len(x) < 2:
            return None, "データが2点未満のため移動範囲を計算できません", 0
        mr_bar = float(np.mean(np.abs(np.diff(x))))
        return mr_bar / CONTROL_CHART_FACTORS[2]['d2'], "MR̄/d2（移動範囲 n=2）", len(x) - 1
    if subgroup_size not in CONTROL_CHART_FACTORS:
        return None, f"サブグループサイズ {subgroup_size} の係数がありません", 0
    k = len(x) // subgroup_size
    if k < 2:
        return None, f"完全なサブグループが {k} 組しかないため群内変動を推定できません（2組以上必要）", k
    groups = x[:k * subgroup_size].reshape(k, subgroup_size)
    r_bar = float(np.mean(groups.max(axis=1) - groups.min(axis=1)))
    return r_bar / CONTROL_CHART_FACTORS[subgroup_size]['d2'], f"R̄/d2（n={subgroup_size}）", k


def capability_indices(mean, sigma, usl, lsl):
    """
    σ と平均から指数を計算する。
    戻り値: (C, Ck, 上側, 下側)。C は両側規格のときだけ。Ck は両側なら min(上側, 下側)、片側ならその側の値
    """
    if sigma is None or sigma <= 0:
        return None, None, None, None
    upper = (usl - mean) / (3 * sigma) if usl is not None else None
    lower = (mean - lsl) / (3 * sigma) if lsl is not None else None
    if usl is not None and lsl is not None:
        return (usl - lsl) / (6 * sigma), min(upper, lower), upper, lower
    return None, (upper if upper is not None else lower), upper, lower


def compute_capability(values, usl, lsl, subgroup_size, ddof=1, alpha=0.05):
    """
    Cp/Cpk（群内）と Pp/Ppk（全体）をまとめて計算する。
    ddof は全体の標準偏差の自由度補正（1: 標本標準偏差 s、0: 母集団標準偏差）。
    戻り値は dict。notes に計算上の注意（ログに出す文）を入れる
    """
    x = np.asarray(values, dtype=float)
    n = len(x)
    notes = []
    mean = float(np.mean(x))
    s_overall = float(np.std(x, ddof=ddof))
    sw, sw_method, k = sigma_within(x, subgroup_size)
    if sw is None:
        notes.append(f"Cp/Cpk を計算できません: {sw_method}")
    else:
        if subgroup_size >= 2:
            left = n - k * subgroup_size
            if left:
                notes.append(f"サブグループに満たない末尾 {left} 点は σ(群内) の推定から除外しました")
            if k < RECOMMENDED_SUBGROUPS:
                notes.append(f"サブグループ数が {k} 組です（工程能力の調査では {RECOMMENDED_SUBGROUPS} 組程度が目安）")
        elif n - 1 < RECOMMENDED_SUBGROUPS:
            notes.append(f"移動範囲が {n - 1} 個です（{RECOMMENDED_SUBGROUPS} 個程度が目安）")
        if sw == 0:
            notes.append("σ(群内) が 0 のため Cp/Cpk を計算できません")
            sw = None
    cp, cpk, cpu, cpl = capability_indices(mean, sw, usl, lsl)
    pp, ppk, ppu, ppl = capability_indices(mean, s_overall, usl, lsl)
    pp_ci = cp_confidence_interval(pp, n, alpha)
    ppk_ci = cpk_confidence_interval_bissell(ppk, n, alpha)
    if cpk is not None and ppk is not None and ppk > 0 and cpk / ppk > 1.33:
        notes.append("Cpk が Ppk を大きく上回っています。サブグループ間の変動（平均のずれ・ドリフト）が大きく、"
                     "工程が統計的管理状態にない可能性があります。管理図を確認してください")
    return {
        "n": n, "mean": mean, "sigma_within": sw, "sigma_within_method": sw_method if sw is not None else None,
        "subgroups": k, "sigma_overall": s_overall,
        "Cp": cp, "Cpk": cpk, "Cpu": cpu, "Cpl": cpl,
        "Pp": pp, "Ppk": ppk, "Ppu": ppu, "Ppl": ppl,
        "Pp_ci": pp_ci, "Ppk_ci": ppk_ci, "notes": notes,
    }

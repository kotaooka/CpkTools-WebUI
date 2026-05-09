import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt
import datetime
import os
import json
import numpy as np
import math
from PIL import Image
import gradio as gr
import re
from contextlib import contextmanager

# =========================================================================
# 定数・設定
# =========================================================================
GITHUB_URL = "https://github.com/kotaooka/CpkTools-WebUI"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 日本語フォント設定 (Windowsの場合)
plt.rcParams['font.family'] = 'MS Gothic'


def get_version():
    try:
        with open("version.txt", "r") as f:
            return f.read().strip()
    except Exception:
        return "バージョン情報未設定"


# -------------------------
# ファイル名サニタイズ関数
# ※ひらがな: \u3040-\u309F, カタカナ: \u30A0-\u30FF, 漢字: \u4E00-\u9FFF, 全角数字: \uFF10-\uFF19 も許可
def sanitize_filename(name):
    name = str(name)
    allowed_pattern = r'[^A-Za-z0-9\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF\uFF10-\uFF19_\-]'
    return re.sub(allowed_pattern, '_', name, flags=re.UNICODE)


# =========================================================================
# 管理図係数 - AIAG / ASTM E2587 準拠
# =========================================================================
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


# =========================================================================
# Cpk / Cp 信頼区間 (Bissell 近似 / カイ二乗ベース)
# =========================================================================
def cpk_confidence_interval_bissell(cpk_estimate, n, alpha=0.05):
    """Bissell (1990) 近似で Cpk の信頼区間を求める。"""
    if n < 2:
        return None, None
    if cpk_estimate == 0:
        return None, None
    z = stats.norm.ppf(1 - alpha / 2)
    se_factor = math.sqrt(1.0 / (9.0 * n * cpk_estimate ** 2) + 1.0 / (2.0 * (n - 1)))
    lcl = cpk_estimate * (1 - z * se_factor)
    ucl = cpk_estimate * (1 + z * se_factor)
    return min(lcl, ucl), max(lcl, ucl)


def cp_confidence_interval(cp_estimate, n, alpha=0.05):
    """Cp の信頼区間 (カイ二乗ベース、Minitab 準拠)"""
    if n < 2:
        return None, None
    df_chi = n - 1
    chi2_lower = stats.chi2.ppf(alpha / 2, df_chi)
    chi2_upper = stats.chi2.ppf(1 - alpha / 2, df_chi)
    lcl = cp_estimate * math.sqrt(chi2_lower / df_chi)
    ucl = cp_estimate * math.sqrt(chi2_upper / df_chi)
    return lcl, ucl


# =========================================================================
# matplotlib figure のリーク対策コンテキストマネージャ
# =========================================================================
@contextmanager
def safe_figure(*args, **kwargs):
    fig = plt.figure(*args, **kwargs)
    try:
        yield fig
    finally:
        plt.close(fig)


# -------------------------
# Excelファイルの読み込み
def read_excel_file(uploaded_file, include_first_row, include_first_column):
    """
    include_first_row が True の場合はすべての行をデータとして扱い、
    False の場合は先頭行をヘッダーとして読み込む。
    include_first_column が True の場合はすべての列をデータとして扱い、
    False の場合は先頭列をインデックスとして読み込む。
    """
    header = None if include_first_row else 0
    index_col = None if include_first_column else 0
    df = pd.read_excel(uploaded_file.name, header=header, index_col=index_col)
    return df


# =========================================================================
# 共通関数: 表示用にラベルを文字列化したDataFrameを返す
#   Excel本来のヘッダー(A, B, C... や 1, 2, 3...)をそのまま使用するため、
#   列名・行名を機械的なラベルに上書きしない。
# =========================================================================
def to_display_df(df):
    """Gradio表示用に列名・行名を文字列化したDataFrameを返す。"""
    df_disp = df.copy()
    df_disp.columns = [str(c) for c in df_disp.columns]
    df_disp.index = [str(i) for i in df_disp.index]
    return df_disp


# =========================================================================
# 共通関数: ターゲットラベル(任意の文字列)からデータ系列を取り出す
#   df.columns / df.index と文字列比較で対応する位置を特定する。
# =========================================================================
def extract_target_series(df, target_label, calc_direction):
    """
    指定されたラベル(列または行のラベル)からデータ系列を取り出す。
    戻り値: (data_series, actual_label, error_message)
    """
    target_str = str(target_label)
    if calc_direction == "列方向":
        cols_str = [str(c) for c in df.columns]
        if target_str not in cols_str:
            return None, None, f"列ラベルが見つかりません: {target_label}"
        idx = cols_str.index(target_str)
        actual_label = str(df.columns[idx])
        data = df.iloc[:, idx]
    else:
        idx_str = [str(i) for i in df.index]
        if target_str not in idx_str:
            return None, None, f"行ラベルが見つかりません: {target_label}"
        idx = idx_str.index(target_str)
        actual_label = str(df.index[idx])
        data = df.iloc[idx]

    # 数値変換できないデータは NaN にして除外
    data = pd.to_numeric(data, errors='coerce')
    return data, actual_label, None


# -------------------------
# 補助関数:選択された列(または行)に合わせて規格値テーブルを更新する
def update_spec_df_with_checkbox(selected_targets, same_spec, current_spec):
    if not selected_targets:
        return []
    if isinstance(current_spec, pd.DataFrame):
        current_spec_list = current_spec.values.tolist()
    elif current_spec is None:
        current_spec_list = []
    else:
        current_spec_list = current_spec

    new_spec = []
    for idx, target in enumerate(selected_targets):
        if idx < len(current_spec_list) and current_spec_list[idx] and str(current_spec_list[idx][0]) == str(target):
            new_spec.append(current_spec_list[idx])
        else:
            new_spec.append([target, "", ""])
    if same_spec and len(new_spec) > 0:
        first_usl = new_spec[0][1]
        first_lsl = new_spec[0][2]
        new_spec = [[row[0], first_usl, first_lsl] for row in new_spec]
    return new_spec


# -------------------------
# ファイルアップロード時のプレビュー更新
def update_preview(uploaded_file, include_first_row, include_first_column, calc_direction):
    if uploaded_file is None:
        return None, gr.update(choices=[])
    try:
        df = read_excel_file(uploaded_file, include_first_row, include_first_column)
    except Exception as e:
        return f"ファイル読み込みエラー: {e}", gr.update(choices=[])
    if df.empty:
        return "ファイルにデータがありません。", gr.update(choices=[])

    # Excel本来のラベルをそのまま使用
    df_display = to_display_df(df)

    # プレビューはインデックス列を可視化するため reset_index で列に変換
    preview_df = df_display.head(5).reset_index().rename(columns={"index": "(行ラベル)"})

    # 解析対象選択肢
    if calc_direction == "列方向":
        target_choices = [str(c) for c in df.columns]
    else:
        target_choices = [str(i) for i in df.index]

    return preview_df, gr.update(choices=target_choices)


# =========================================================================
# Cpk解析結果 HTML サマリー生成
# =========================================================================
def generate_cpk_html(results):
    if not results:
        return ""

    def _color(cpk):
        if cpk is None:
            return "#ffffff"
        if cpk >= 1.67:
            return "#c8e6c9"
        if cpk >= 1.33:
            return "#dcedc8"
        if cpk >= 1.0:
            return "#fff9c4"
        return "#ffcdd2"

    def _label(cpk):
        if cpk is None:
            return "―"
        if cpk >= 1.67:
            return "◎ 非常に良好"
        if cpk >= 1.33:
            return "○ 良好"
        if cpk >= 1.0:
            return "△ 要注意"
        return "✕ 不良"

    rows = ""
    for r in results:
        cpk = r.get("Cpk")
        cp = r.get("Cp")
        cpk_lo = r.get("Cpk_lower (95%, Bissell)")
        cpk_hi = r.get("Cpk_upper (95%, Bissell)")
        ci = (f"{cpk_lo:.3f} ～ {cpk_hi:.3f}"
              if cpk_lo is not None and cpk_hi is not None else "―")
        bg = _color(cpk)
        cell = "padding:5px 10px;border:1px solid #aaa;color:#222222"
        rows += (
            f'<tr style="background:{bg};color:#222222">'
            f'<td style="{cell}">{r["解析対象"]}</td>'
            f'<td style="{cell};text-align:center">{r["サンプル数"]}</td>'
            f'<td style="{cell};text-align:center">{r["平均値"]:.4f}</td>'
            f'<td style="{cell};text-align:center">{r["標準偏差"]:.4f}</td>'
            f'<td style="{cell};text-align:center">{"―" if cp is None else f"{cp:.3f}"}</td>'
            f'<td style="{cell};text-align:center;font-weight:bold">{"―" if cpk is None else f"{cpk:.3f}"}</td>'
            f'<td style="{cell};text-align:center">{ci}</td>'
            f'<td style="{cell};text-align:center;font-weight:bold">{_label(cpk)}</td>'
            f'</tr>\n'
        )

    th = "padding:5px 10px;border:1px solid #aaa;color:#222222;background:#dddddd"
    legend = (
        '<div style="margin-top:6px;font-size:0.82em;display:flex;gap:8px;flex-wrap:wrap">'
        '<span style="background:#c8e6c9;color:#222222;padding:2px 8px;border-radius:3px">◎ Cpk ≥ 1.67</span>'
        '<span style="background:#dcedc8;color:#222222;padding:2px 8px;border-radius:3px">○ 1.33 ≤ Cpk &lt; 1.67</span>'
        '<span style="background:#fff9c4;color:#222222;padding:2px 8px;border-radius:3px">△ 1.00 ≤ Cpk &lt; 1.33</span>'
        '<span style="background:#ffcdd2;color:#222222;padding:2px 8px;border-radius:3px">✕ Cpk &lt; 1.00</span>'
        '</div>'
    )

    return (
        '<table style="border-collapse:collapse;width:100%;font-size:0.9em">'
        '<thead>'
        f'<tr><th style="{th};text-align:left">解析対象</th>'
        f'<th style="{th}">n</th>'
        f'<th style="{th}">平均値</th>'
        f'<th style="{th}">標準偏差</th>'
        f'<th style="{th}">Cp</th>'
        f'<th style="{th}">Cpk</th>'
        f'<th style="{th}">Cpk 95%CI</th>'
        f'<th style="{th}">判定</th>'
        '</tr></thead>'
        f'<tbody>\n{rows}</tbody>'
        '</table>'
        f'{legend}'
    )


# =========================================================================
# 解析処理(工程能力解析ツール)
# =========================================================================
def run_analysis(uploaded_file, selected_targets, spec_table, subgroup_size, include_first_row, include_first_column,
                 calc_direction, show_hist, show_qq, show_density, show_xbar, show_r, show_s, std_method):
    log_messages = ""
    hist_images = []
    qq_images = []
    density_images = []
    xbar_images = []
    r_images = []
    s_images = []
    excel_file = None
    excel_preview = None
    results = []
    cpk_html = ""

    ddof_value = 1 if std_method == "サンプル標準偏差" else 0

    if subgroup_size < 1 or subgroup_size > 10:
        return (f"エラー: サブグループサイズは1～10の範囲で指定してください(指定値: {subgroup_size})",
                "", None, None, None, None, None, None, None, None)

    if uploaded_file is None:
        return "エラー: ファイルが選択されていません", "", None, None, None, None, None, None, None, None
    try:
        df = read_excel_file(uploaded_file, include_first_row, include_first_column)
        log_messages += "ファイル読み込み成功。\n"
    except Exception as e:
        return f"エラー: ファイル読み込み中に問題が発生しました: {e}", "", None, None, None, None, None, None, None, None

    if df.empty:
        return "エラー: ファイルにデータがありません", "", None, None, None, None, None, None, None, None
    if not selected_targets:
        return "エラー: 解析対象が選択されていません", "", None, None, None, None, None, None, None, None
    try:
        spec_df = pd.DataFrame(spec_table, columns=["解析対象", "規格上限値", "規格下限値"])
    except Exception as e:
        return f"エラー: 規格値テーブルの読み込みに失敗しました: {e}", "", None, None, None, None, None, None, None, None

    if len(spec_df) != len(selected_targets):
        return "エラー: 選択された対象数と規格値入力の行数が一致しません。", "", None, None, None, None, None, None, None, None

    # 規格値未入力の事前チェック（両方空欄の行を警告）
    for i in range(len(selected_targets)):
        try:
            usl_raw = str(spec_df.iloc[i]["規格上限値"]).strip()
            lsl_raw = str(spec_df.iloc[i]["規格下限値"]).strip()
            if usl_raw == "" and lsl_raw == "":
                log_messages += f"⚠ 警告: 「{selected_targets[i]}」の規格値(上限・下限)が両方未入力です。スキップされます。\n"
        except Exception:
            pass

    run_timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')

    for i, target_label in enumerate(selected_targets):
        data, actual_label, err = extract_target_series(df, target_label, calc_direction)
        if err is not None:
            log_messages += f"エラー: {err}\n"
            continue

        # 欠損値除去
        n_before = len(data)
        data = data.dropna()
        sample_n = len(data)
        if sample_n < n_before:
            log_messages += (f"注意: {target_label} で {n_before - sample_n} 件の欠損値/非数値を除外しました"
                             f"(元: {n_before} → 解析: {sample_n})。\n")
        if sample_n == 0:
            log_messages += f"エラー: {target_label} に有効な数値データがありません。\n"
            continue
        if sample_n < 2:
            log_messages += f"エラー: {target_label} のサンプル数が不足しています(n={sample_n})。\n"
            continue

        # 規格値読み込み
        try:
            spec_entry_usl = spec_df.iloc[i]["規格上限値"]
            spec_entry_lsl = spec_df.iloc[i]["規格下限値"]
            current_usl = float(spec_entry_usl) if str(spec_entry_usl).strip() != "" else None
            current_lsl = float(spec_entry_lsl) if str(spec_entry_lsl).strip() != "" else None
        except Exception as e:
            log_messages += f"エラー: {target_label} の規格値が正しく入力されていません: {e}\n"
            continue

        if current_usl is not None and current_lsl is not None and current_usl <= current_lsl:
            log_messages += f"エラー: {target_label} の規格上限値({current_usl})が下限値({current_lsl})以下です。\n"
            continue

        try:
            max_val = float(data.max())
            min_val = float(data.min())
            std_val = float(data.std(ddof=ddof_value))
            mean_val = float(data.mean())
            kurtosis_val = float(data.kurtosis())
            skewness_val = float(data.skew())
            if std_val == 0:
                log_messages += f"エラー: {target_label} の標準偏差が0のため、Cp/Cpk計算をスキップ(n={sample_n})。\n"
                continue

            # 正規性検定 (Shapiro-Wilk)
            if 3 <= sample_n <= 5000:
                try:
                    sw_stat, sw_p = stats.shapiro(data)
                    if sw_p < 0.05:
                        log_messages += (f"警告: {target_label} はShapiro-Wilk検定でp={sw_p:.4f}<0.05、"
                                         f"正規性に疑義あり。Cp/Cpk値の解釈に注意。\n")
                except Exception:
                    pass

            # === Cp / Cpk の点推定 ===
            if current_usl is not None and current_lsl is not None:
                spec_type = "両側"
                Cp = (current_usl - current_lsl) / (6 * std_val)
                Cpk = min((current_usl - mean_val), (mean_val - current_lsl)) / (3 * std_val)
            elif current_usl is not None:
                spec_type = "上側のみ"
                Cp = (current_usl - mean_val) / (3 * std_val)
                Cpk = Cp
            elif current_lsl is not None:
                spec_type = "下側のみ"
                Cp = (mean_val - current_lsl) / (3 * std_val)
                Cpk = Cp
            else:
                log_messages += f"エラー: {target_label} の規格値が入力されていません。\n"
                continue

            Cp_lower, Cp_upper = cp_confidence_interval(Cp, sample_n, alpha=0.05)
            Cpk_lower, Cpk_upper = cpk_confidence_interval_bissell(Cpk, sample_n, alpha=0.05)

            results.append({
                "解析対象": target_label,
                "サンプル数": sample_n,
                "規格種別": spec_type,
                "上限規格": current_usl,
                "下限規格": current_lsl,
                "最大値": max_val,
                "最小値": min_val,
                "標準偏差": std_val,
                "平均値": mean_val,
                "Cp": Cp,
                "Cp_lower (95%)": Cp_lower,
                "Cp_upper (95%)": Cp_upper,
                "Cpk": Cpk,
                "Cpk_lower (95%, Bissell)": Cpk_lower,
                "Cpk_upper (95%, Bissell)": Cpk_upper,
                "尖度": kurtosis_val,
                "歪度": skewness_val
            })
            log_messages += f"解析対象: {target_label} ({actual_label}) の統計計算完了(n={sample_n})。\n"
        except Exception as e:
            log_messages += f"エラー: {target_label} の統計計算中に問題が発生しました: {e}\n"
            continue

        safe_col_label = sanitize_filename(target_label)

        # === ヒストグラム生成 ===
        if show_hist:
            try:
                with safe_figure() as fig:
                    bins = max(5, int(math.ceil(math.log2(sample_n) + 1)))
                    plt.hist(data, bins=bins, color="skyblue", edgecolor="black", label="度数")
                    plt.axvline(mean_val, color="orange", linestyle="-", linewidth=1.5, label="平均値")
                    if current_usl is not None:
                        plt.axvline(current_usl, color="green", linestyle="-.", linewidth=1.5, label="USL")
                    if current_lsl is not None:
                        plt.axvline(current_lsl, color="purple", linestyle="-.", linewidth=1.5, label="LSL")
                    plt.xlabel("値")
                    plt.ylabel("度数")
                    plt.title(f"ヒストグラム ({target_label})")
                    plt.legend(fontsize=8)
                    hist_filename = os.path.join(OUTPUT_DIR, f"{run_timestamp}_hist_{safe_col_label}.jpg")
                    plt.savefig(hist_filename, format="jpg", bbox_inches='tight')
                    hist_images.append(hist_filename)
                log_messages += f"{target_label} のヒストグラム生成完了。\n"
            except Exception as e:
                log_messages += f"エラー: {target_label} のヒストグラム生成中: {e}\n"

        # === QQプロット生成 ===
        if show_qq:
            try:
                with safe_figure() as fig:
                    stats.probplot(data, dist="norm", plot=plt)
                    plt.title(f"QQプロット ({target_label})")
                    qq_filename = os.path.join(OUTPUT_DIR, f"{run_timestamp}_qq_{safe_col_label}.jpg")
                    plt.savefig(qq_filename, format="jpg", bbox_inches='tight')
                    qq_images.append(qq_filename)
                log_messages += f"{target_label} のQQプロット生成完了。\n"
            except Exception as e:
                log_messages += f"エラー: {target_label} のQQプロット生成中: {e}\n"

        # === 確率密度分布生成 ===
        if show_density:
            try:
                with safe_figure() as fig:
                    x_vals = np.linspace(mean_val - 4 * std_val, mean_val + 4 * std_val, 200)
                    y_vals = stats.norm.pdf(x_vals, loc=mean_val, scale=std_val)
                    plt.plot(x_vals, y_vals, label="正規分布", color="blue")
                    plt.axvline(mean_val - 3 * std_val, color="red", linestyle="--", label="-3σ")
                    plt.axvline(mean_val + 3 * std_val, color="red", linestyle="--", label="+3σ")
                    plt.axvline(mean_val, color="orange", linestyle="-", label="平均値")
                    if current_usl is not None:
                        plt.axvline(current_usl, color="green", linestyle="-.", label="規格上限値")
                    if current_lsl is not None:
                        plt.axvline(current_lsl, color="purple", linestyle="-.", label="規格下限値")
                    ax = plt.gca()
                    y_top = ax.get_ylim()[1]
                    label_y = y_top * 0.20
                    ax.text(mean_val - 3 * std_val, label_y, f"-3σ: {mean_val - 3 * std_val:.2f}",
                            rotation=90, color="black", ha="center", va="bottom", fontsize=8)
                    ax.text(mean_val + 3 * std_val, label_y, f"+3σ: {mean_val + 3 * std_val:.2f}",
                            rotation=90, color="black", ha="center", va="bottom", fontsize=8)
                    ax.text(mean_val, label_y, f"平均値: {mean_val:.2f}",
                            rotation=90, color="black", ha="center", va="bottom", fontsize=8)
                    if current_usl is not None:
                        ax.text(current_usl, label_y, f"USL: {current_usl:.2f}",
                                rotation=90, color="black", ha="center", va="bottom", fontsize=8)
                    if current_lsl is not None:
                        ax.text(current_lsl, label_y, f"LSL: {current_lsl:.2f}",
                                rotation=90, color="black", ha="center", va="bottom", fontsize=8)
                    plt.xlabel("値")
                    plt.ylabel("確率密度")
                    plt.title(f"確率密度分布 ({target_label})")
                    plt.legend()
                    density_filename = os.path.join(OUTPUT_DIR, f"{run_timestamp}_density_{safe_col_label}.jpg")
                    plt.savefig(density_filename, format="jpg", bbox_inches='tight')
                    density_images.append(density_filename)
                log_messages += f"{target_label} の確率密度分布描画完了。\n"
            except Exception as e:
                log_messages += f"エラー: {target_label} の確率密度分布描画中: {e}\n"

        # === 管理図 ===
        if subgroup_size == 1:
            individuals = data.values
            n_individuals = len(individuals)
            if n_individuals < 2:
                log_messages += f"警告: {target_label} のデータ点数が不足、I/MR管理図を生成できません。\n"
            else:
                i_bar = float(np.mean(individuals))
                moving_ranges = np.abs(np.diff(individuals))
                mr_bar = float(np.mean(moving_ranges))
                sigma_short = mr_bar / CONTROL_CHART_FACTORS[2]['d2']
                UCL_i = i_bar + 3 * sigma_short
                LCL_i = i_bar - 3 * sigma_short

                if show_xbar:
                    try:
                        with safe_figure() as fig:
                            plt.plot(range(1, n_individuals + 1), individuals, marker='o',
                                     linestyle='-', color='blue', label='値')
                            plt.axhline(i_bar, color='green', linestyle='--', label='平均')
                            plt.axhline(UCL_i, color='red', linestyle='--', label=f'UCL={UCL_i:.3f}')
                            plt.axhline(LCL_i, color='red', linestyle='--', label=f'LCL={LCL_i:.3f}')
                            plt.xlabel('データ点')
                            plt.ylabel('値')
                            plt.title(f"I管理図 ({target_label})")
                            plt.legend(fontsize=8)
                            i_filename = os.path.join(OUTPUT_DIR, f"{run_timestamp}_i_{safe_col_label}.jpg")
                            plt.savefig(i_filename, format="jpg", bbox_inches='tight')
                            xbar_images.append(i_filename)
                        log_messages += f"{target_label} のI管理図生成完了。\n"
                    except Exception as e:
                        log_messages += f"エラー: {target_label} のI管理図生成中: {e}\n"

                if show_r:
                    try:
                        with safe_figure() as fig:
                            UCL_mr = mr_bar * CONTROL_CHART_FACTORS[2]['D4']
                            plt.plot(range(2, n_individuals + 1), moving_ranges, marker='o',
                                     linestyle='-', color='blue', label='移動範囲')
                            plt.axhline(mr_bar, color='green', linestyle='--', label=f'MR-bar={mr_bar:.3f}')
                            plt.axhline(UCL_mr, color='red', linestyle='--', label=f'UCL={UCL_mr:.3f}')
                            plt.axhline(0, color='red', linestyle='--', label='LCL=0')
                            plt.xlabel('データ点 (2番目以降)')
                            plt.ylabel('移動範囲')
                            plt.title(f"MR管理図 ({target_label})")
                            plt.legend(fontsize=8)
                            mr_filename = os.path.join(OUTPUT_DIR, f"{run_timestamp}_mr_{safe_col_label}.jpg")
                            plt.savefig(mr_filename, format="jpg", bbox_inches='tight')
                            r_images.append(mr_filename)
                        log_messages += f"{target_label} のMR管理図生成完了。\n"
                    except Exception as e:
                        log_messages += f"エラー: {target_label} のMR管理図生成中: {e}\n"

                if show_s:
                    log_messages += "情報: サブグループサイズ=1のためs管理図はスキップしました。\n"
        else:
            data_array = data.values
            n_full_groups = len(data_array) // subgroup_size
            if n_full_groups == 0:
                log_messages += f"警告: {target_label} のデータ数(n={sample_n})がサブグループサイズ未満で管理図不可。\n"
            else:
                full_groups = [data_array[k * subgroup_size:(k + 1) * subgroup_size]
                               for k in range(n_full_groups)]
                group_means = np.array([np.mean(g) for g in full_groups])
                group_ranges = np.array([np.max(g) - np.min(g) for g in full_groups])
                group_stds = np.array([np.std(g, ddof=1) for g in full_groups])

                Xbar_center = float(np.mean(group_means))
                Rbar = float(np.mean(group_ranges))
                sbar = float(np.mean(group_stds))
                current_factors = CONTROL_CHART_FACTORS[subgroup_size]

                if show_xbar:
                    try:
                        with safe_figure() as fig:
                            UCL_xbar = Xbar_center + current_factors['A2'] * Rbar
                            LCL_xbar = Xbar_center - current_factors['A2'] * Rbar
                            plt.plot(range(1, len(group_means) + 1), group_means,
                                     marker='o', linestyle='-', color='blue', label='サブグループ平均')
                            plt.axhline(Xbar_center, color='green', linestyle='--', label='全体平均')
                            plt.axhline(UCL_xbar, color='red', linestyle='--', label=f'UCL={UCL_xbar:.3f}')
                            plt.axhline(LCL_xbar, color='red', linestyle='--', label=f'LCL={LCL_xbar:.3f}')
                            plt.xlabel('サブグループ番号')
                            plt.ylabel('サブグループ平均')
                            plt.title(f"X-bar管理図 ({target_label})")
                            plt.legend(fontsize=8)
                            xbar_filename = os.path.join(OUTPUT_DIR, f"{run_timestamp}_xbar_{safe_col_label}.jpg")
                            plt.savefig(xbar_filename, format="jpg", bbox_inches='tight')
                            xbar_images.append(xbar_filename)
                        log_messages += f"{target_label} のX-bar管理図生成完了。\n"
                    except Exception as e:
                        log_messages += f"エラー: {target_label} のX-bar管理図生成中: {e}\n"

                if show_r:
                    try:
                        with safe_figure() as fig:
                            UCL_R = current_factors['D4'] * Rbar
                            LCL_R = current_factors['D3'] * Rbar
                            plt.plot(range(1, len(group_ranges) + 1), group_ranges,
                                     marker='o', linestyle='-', color='blue', label='サブグループレンジ')
                            plt.axhline(Rbar, color='green', linestyle='--', label=f'R-bar={Rbar:.3f}')
                            plt.axhline(UCL_R, color='red', linestyle='--', label=f'UCL={UCL_R:.3f}')
                            plt.axhline(LCL_R, color='red', linestyle='--', label=f'LCL={LCL_R:.3f}')
                            plt.xlabel('サブグループ番号')
                            plt.ylabel('レンジ')
                            plt.title(f"R管理図 ({target_label})")
                            plt.legend(fontsize=8)
                            r_filename = os.path.join(OUTPUT_DIR, f"{run_timestamp}_r_{safe_col_label}.jpg")
                            plt.savefig(r_filename, format="jpg", bbox_inches='tight')
                            r_images.append(r_filename)
                        log_messages += f"{target_label} のR管理図生成完了。\n"
                    except Exception as e:
                        log_messages += f"エラー: {target_label} のR管理図生成中: {e}\n"

                if show_s:
                    try:
                        with safe_figure() as fig:
                            UCL_s = sbar * current_factors['B4']
                            LCL_s = sbar * current_factors['B3']
                            plt.plot(range(1, len(group_stds) + 1), group_stds,
                                     marker='o', linestyle='-', color='blue', label='サブグループ標準偏差')
                            plt.axhline(sbar, color='green', linestyle='--', label=f's-bar={sbar:.3f}')
                            plt.axhline(UCL_s, color='red', linestyle='--', label=f'UCL={UCL_s:.3f}')
                            plt.axhline(LCL_s, color='red', linestyle='--', label=f'LCL={LCL_s:.3f}')
                            plt.xlabel('サブグループ番号')
                            plt.ylabel('標準偏差')
                            plt.title(f"s管理図 ({target_label})")
                            plt.legend(fontsize=8)
                            s_filename = os.path.join(OUTPUT_DIR, f"{run_timestamp}_s_{safe_col_label}.jpg")
                            plt.savefig(s_filename, format="jpg", bbox_inches='tight')
                            s_images.append(s_filename)
                        log_messages += f"{target_label} のs管理図生成完了。\n"
                    except Exception as e:
                        log_messages += f"エラー: {target_label} のs管理図生成中: {e}\n"

    if results:
        cpk_html = generate_cpk_html(results)
        output_filename = os.path.join(OUTPUT_DIR, f"{run_timestamp}_results.xlsx")
        try:
            results_df = pd.DataFrame(results)
            results_df.to_excel(output_filename, index=False)
            excel_file = output_filename
            excel_preview = results_df
            log_messages += f"Excelファイル出力完了: {output_filename}\n"
        except Exception as e:
            log_messages += f"エラー: Excelファイル書き出し中: {e}\n"
    else:
        log_messages += "エラー: 解析対象から有効なデータが得られませんでした。\n"

    return log_messages, cpk_html, hist_images, qq_images, density_images, xbar_images, r_images, s_images, excel_file, excel_preview


# =========================================================================
# F検定 / t検定 / 相関
# =========================================================================
def run_stat_test(uploaded_file, selected_targets, perform_f_test, alpha_f, perform_t_test, ttest_variant,
                  alpha_t, include_first_row, include_first_column, plot_overlay, calc_corr, calc_direction):
    log_messages = ""
    density_images = []
    excel_file = None
    excel_preview = None

    t_test_done = False
    f_test_done = False
    t_stat, p_value_t, df_t = None, None, None
    f_stat, p_value_f, dfn, dfd = None, None, None, None
    significance_f = None
    significance_t = None

    if uploaded_file is None:
        return "エラー: ファイルが選択されていません", [], None, None

    try:
        df = read_excel_file(uploaded_file, include_first_row, include_first_column)
        log_messages += "ファイル読み込み成功。\n"
    except Exception as e:
        return f"エラー: ファイル読み込み中: {e}", [], None, None

    if df.empty:
        return "エラー: ファイルにデータがありません", [], None, None

    if len(selected_targets) != 2:
        return "エラー: 検定対象は2つ選択してください", [], None, None

    raw_data_list = []
    col_names = []
    for target_label in selected_targets:
        data, actual_label, err = extract_target_series(df, target_label, calc_direction)
        if err is not None:
            return f"エラー: {err}", [], None, None
        col_names.append(target_label)
        raw_data_list.append(data)

    if len(raw_data_list) != 2:
        return "エラー: 選択された2対象のデータに問題があります。", [], None, None

    # 対応ありt検定では行インデックスでペアリングしてからdropna
    is_paired = (perform_t_test != "T検定を実施しない" and ttest_variant == "対応ありt検定")
    if is_paired:
        df_pair = pd.DataFrame({
            "data1": raw_data_list[0].values,
            "data2": raw_data_list[1].values
        }).dropna()
        if len(df_pair) < 2:
            return "エラー: 対応ありt検定に必要な有効ペアが2組未満です。", [], None, None
        data1 = df_pair["data1"].reset_index(drop=True)
        data2 = df_pair["data2"].reset_index(drop=True)
        log_messages += f"情報: 対応ありt検定のため有効ペア{len(df_pair)}組で解析します。\n"
    else:
        data1 = raw_data_list[0].dropna().reset_index(drop=True)
        data2 = raw_data_list[1].dropna().reset_index(drop=True)

    if len(data1) == 0 or len(data2) == 0:
        return "エラー: いずれかのデータが空です。", [], None, None

    n1, n2 = len(data1), len(data2)
    mean1, mean2 = float(data1.mean()), float(data2.mean())
    var1, var2 = float(data1.var(ddof=1)), float(data2.var(ddof=1))

    # ----- F検定 -----
    if perform_f_test == "F検定を実施する":
        if var2 == 0 or var1 == 0:
            log_messages += "警告: 分散が0のためF検定をスキップします。\n"
        else:
            if var1 >= var2:
                f_stat = var1 / var2
                dfn = n1 - 1
                dfd = n2 - 1
            else:
                f_stat = var2 / var1
                dfn = n2 - 1
                dfd = n1 - 1
            p_value_f = 2 * stats.f.sf(f_stat, dfn, dfd)
            significance_f = "有意差あり" if p_value_f < alpha_f else "有意差なし"
            log_messages += (
                f"F検定結果:\n"
                f"  サンプル数: {n1}, {n2}\n"
                f"  平均: {mean1:.4f}, {mean2:.4f}\n"
                f"  分散: {var1:.4f}, {var2:.4f}\n"
                f"  F値: {f_stat:.4f}\n"
                f"  自由度: ({dfn}, {dfd})\n"
                f"  P値: {p_value_f:.4f}\n"
                f"  判定 (有意水準 {alpha_f}): {significance_f}\n"
            )
            f_test_done = True

    # ----- t検定 -----
    if perform_t_test != "T検定を実施しない":
        try:
            if ttest_variant == "対応ありt検定":
                t_stat, p_value_t = stats.ttest_rel(data1, data2)
                df_t = n1 - 1
            elif ttest_variant == "独立t検定(分散が等しい)":
                t_stat, p_value_t = stats.ttest_ind(data1, data2, equal_var=True)
                df_t = n1 + n2 - 2
            elif ttest_variant == "独立t検定(分散が異なる)":
                t_stat, p_value_t = stats.ttest_ind(data1, data2, equal_var=False)
                if var1 > 0 and var2 > 0:
                    df_t = ((var1 / n1 + var2 / n2) ** 2 /
                            ((var1 / n1) ** 2 / (n1 - 1) + (var2 / n2) ** 2 / (n2 - 1)))
                else:
                    df_t = n1 + n2 - 2
            else:
                return f"エラー: t検定の種類が不正: {ttest_variant}", [], None, None

            significance_t = "有意差あり" if p_value_t < alpha_t else "有意差なし"
            log_messages += (
                f"t検定結果 ({ttest_variant}):\n"
                f"  サンプル数: {n1}, {n2}\n"
                f"  平均: {mean1:.4f}, {mean2:.4f}\n"
                f"  分散: {var1:.4f}, {var2:.4f}\n"
                f"  t値: {t_stat:.4f}\n"
                f"  自由度: {df_t:.2f}\n"
                f"  P値: {p_value_t:.4f}\n"
                f"  判定 (有意水準 {alpha_t}): {significance_t}\n"
            )
            t_test_done = True
        except Exception as e:
            log_messages += f"エラー: t検定中に問題が発生しました: {e}\n"

    result_dict = {
        "検定対象1": col_names[0], "検定対象2": col_names[1],
        "サンプル数1": n1, "サンプル数2": n2,
        "平均1": mean1, "平均2": mean2,
        "分散1": var1, "分散2": var2
    }
    if f_test_done:
        result_dict.update({
            "F値": f_stat,
            "P値 (F検定)": p_value_f,
            "判定 (F検定)": significance_f
        })
    if t_test_done:
        result_dict.update({
            "t値": t_stat,
            "自由度 (t検定)": df_t,
            "P値 (t検定)": p_value_t,
            "判定 (t検定)": significance_t
        })

    # ----- 相関計算 -----
    df_corr = None
    corr_coeff = None
    r2_value = None
    if calc_corr == "相関計算を実施する":
        try:
            df_corr = pd.DataFrame({"data1": data1.values, "data2": data2.values}).dropna()
            if len(df_corr) > 1:
                corr_coeff = float(df_corr["data1"].corr(df_corr["data2"]))
                r2_value = corr_coeff ** 2
                log_messages += f"相関係数: {corr_coeff:.4f}\n決定係数: {r2_value:.4f}\n"
            else:
                log_messages += "警告: 相関計算に十分なデータがありません。\n"
                df_corr = None
        except Exception as e:
            log_messages += f"エラー: 相関計算中: {e}\n"
            df_corr = None

        if corr_coeff is not None:
            result_dict.update({"相関係数": corr_coeff, "決定係数": r2_value})

        if df_corr is not None and len(df_corr) >= 2:
            try:
                with safe_figure() as fig:
                    plt.scatter(df_corr["data1"], df_corr["data2"], color="blue", label="データポイント")
                    slope, intercept = np.polyfit(df_corr["data1"], df_corr["data2"], 1)
                    x_lin = np.linspace(df_corr["data1"].min(), df_corr["data1"].max(), 100)
                    y_lin = slope * x_lin + intercept
                    plt.plot(x_lin, y_lin, color="red", label=f"y={slope:.3f}x+{intercept:.3f}")
                    plt.xlabel(str(selected_targets[0]))
                    plt.ylabel(str(selected_targets[1]))
                    plt.title(f"散布図 (r={corr_coeff:.3f}, R²={r2_value:.3f})")
                    plt.legend(fontsize=8)
                    dt_scatter = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                    scatter_filename = os.path.join(OUTPUT_DIR, f"{dt_scatter}_scatter.jpg")
                    plt.savefig(scatter_filename, format="jpg", bbox_inches='tight')
                    density_images.append(scatter_filename)
                log_messages += "散布図生成完了。\n"
            except Exception as e:
                log_messages += f"エラー: 散布図生成中: {e}\n"

    results_df = pd.DataFrame([result_dict])
    dt_now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_filename = os.path.join(OUTPUT_DIR, f"{dt_now}_stat_test_results.xlsx")
    try:
        results_df.to_excel(excel_filename, index=False)
        excel_file = excel_filename
        excel_preview = results_df
    except Exception as e:
        log_messages += f"エラー: Excelファイル書き出し中: {e}\n"

    # ----- 正規分布の重ね描き -----
    if plot_overlay == "正規分布を表示する":
        try:
            with safe_figure() as fig:
                overall_min = min(data1.min(), data2.min())
                overall_max = max(data1.max(), data2.max())
                range_padding = (overall_max - overall_min) * 0.1
                x_vals = np.linspace(overall_min - range_padding, overall_max + range_padding, 200)
                std1 = float(data1.std(ddof=1))
                std2 = float(data2.std(ddof=1))
                if std1 > 0 and std2 > 0:
                    y1_vals = stats.norm.pdf(x_vals, loc=mean1, scale=std1)
                    y2_vals = stats.norm.pdf(x_vals, loc=mean2, scale=std2)
                    plt.plot(x_vals, y1_vals, label=f"{selected_targets[0]} (μ={mean1:.2f}, σ={std1:.2f})", color="blue")
                    plt.plot(x_vals, y2_vals, label=f"{selected_targets[1]} (μ={mean2:.2f}, σ={std2:.2f})", color="red")
                    plt.xlabel("値")
                    plt.ylabel("確率密度")
                    plt.title("各群正規分布の重ね描き")
                    plt.legend(fontsize=8)
                    dt_now2 = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                    overlay_filename = os.path.join(OUTPUT_DIR, f"{dt_now2}_density_overlay.jpg")
                    plt.savefig(overlay_filename, format="jpg", bbox_inches='tight')
                    density_images.append(overlay_filename)
                    log_messages += "各群正規分布の重ね描きプロット生成完了。\n"
                else:
                    log_messages += "警告: 標準偏差が0のため正規分布重ね描きをスキップ。\n"
        except Exception as e:
            log_messages += f"エラー: 正規分布重ね描き中: {e}\n"

    # ----- t分布プロット -----
    if t_test_done:
        try:
            with safe_figure() as fig:
                lower_bound = stats.t.ppf(0.001, df_t)
                upper_bound = stats.t.ppf(0.999, df_t)
                lower_bound = min(lower_bound, t_stat * 1.1)
                upper_bound = max(upper_bound, t_stat * 1.1)
                x_t = np.linspace(lower_bound, upper_bound, 300)
                y_t = stats.t.pdf(x_t, df_t)
                plt.plot(x_t, y_t, label=f"t分布 (df={df_t:.2f})", color="purple")
                plt.axvline(t_stat, color="black", linestyle="--", label=f"t値 = {t_stat:.2f}")
                plt.text(t_stat, max(y_t) * 0.7, f"p={p_value_t:.4f}", color="black", fontsize=10,
                         rotation=90, ha="left", va="center")
                plt.xlabel("t値")
                plt.ylabel("確率密度")
                plt.title("t分布プロット")
                plt.legend()
                timestamp2 = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                t_plot_filename = os.path.join(OUTPUT_DIR, f"{timestamp2}_t_distribution.jpg")
                plt.savefig(t_plot_filename, format="jpg", bbox_inches='tight')
                density_images.append(t_plot_filename)
            log_messages += "t分布プロット生成完了。\n"
        except Exception as e:
            log_messages += f"エラー: t分布プロット中: {e}\n"

    # ----- F分布プロット -----
    if f_test_done:
        try:
            with safe_figure() as fig:
                lower_bound = stats.f.ppf(0.001, dfn, dfd)
                upper_bound = stats.f.ppf(0.999, dfn, dfd)
                upper_bound = max(upper_bound, f_stat * 1.1)
                x_f = np.linspace(lower_bound, upper_bound, 300)
                y_f = stats.f.pdf(x_f, dfn, dfd)
                plt.plot(x_f, y_f, label=f"F分布 (dfn={dfn}, dfd={dfd})", color="orange")
                plt.axvline(f_stat, color="black", linestyle="--", label=f"F値 = {f_stat:.2f}")
                plt.text(f_stat, max(y_f) * 0.7, f"p={p_value_f:.4f}", color="black", fontsize=10,
                         rotation=90, ha="left", va="center")
                plt.xlabel("F値")
                plt.ylabel("確率密度")
                plt.title("F分布プロット")
                plt.legend()
                timestamp3 = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                f_plot_filename = os.path.join(OUTPUT_DIR, f"{timestamp3}_f_distribution.jpg")
                plt.savefig(f_plot_filename, format="jpg", bbox_inches='tight')
                density_images.append(f_plot_filename)
            log_messages += "F分布プロット生成完了。\n"
        except Exception as e:
            log_messages += f"エラー: F分布プロット中: {e}\n"

    return log_messages, density_images, excel_file, excel_preview


# -------------------------
# Outputフォルダを開く関数(Windows専用)
def open_output_folder():
    folder_path = os.path.abspath(OUTPUT_DIR)
    try:
        os.startfile(folder_path)
    except Exception as e:
        print(f"フォルダを開くのに失敗しました: {e}")


# =========================================================================
# 設定 JSON のエクスポート / インポート
# =========================================================================
SETTINGS_SCHEMA_VERSION = "1.0"


def export_settings_to_json(subgroup_size, std_method,
                            include_first_row, include_first_column, calc_direction,
                            show_hist, show_qq, show_density,
                            show_xbar, show_r, show_s,
                            same_spec, selected_targets, spec_table):
    """現在のUI設定をJSONファイルとして出力する。"""
    if isinstance(spec_table, pd.DataFrame):
        spec_list = spec_table.values.tolist()
    elif spec_table is None:
        spec_list = []
    else:
        spec_list = list(spec_table)

    # 値の型を JSON 互換に変換
    spec_list_clean = []
    for row in spec_list:
        if row is None:
            continue
        spec_list_clean.append([str(x) if x is not None else "" for x in row])

    settings = {
        "schema_version": SETTINGS_SCHEMA_VERSION,
        "exported_at": datetime.datetime.now().isoformat(),
        "subgroup_size": int(subgroup_size) if subgroup_size is not None else 5,
        "std_method": std_method,
        "include_first_row": bool(include_first_row),
        "include_first_column": bool(include_first_column),
        "calc_direction": calc_direction,
        "show_hist": bool(show_hist),
        "show_qq": bool(show_qq),
        "show_density": bool(show_density),
        "show_xbar": bool(show_xbar),
        "show_r": bool(show_r),
        "show_s": bool(show_s),
        "same_spec": bool(same_spec),
        "selected_targets": [str(t) for t in (selected_targets or [])],
        "spec_table": spec_list_clean,
    }

    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = os.path.join(OUTPUT_DIR, f"{timestamp}_settings.json")
    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception as e:
        return None, f"エラー: 設定JSONの書き出しに失敗しました: {e}"
    return filename, f"設定を保存しました: {filename}"


def import_settings_from_json(json_file):
    """
    JSONファイルから設定を読み込み、UI各コンポーネントの更新値を返す。
    戻り値は順序固定で:
      (subgroup_size, std_method,
       include_first_row, include_first_column, calc_direction,
       show_hist, show_qq, show_density,
       show_xbar, show_r, show_s,
       same_spec, selected_targets_dropdown, spec_table, log_message)
    """
    no_change_15 = [gr.update() for _ in range(14)] + [""]

    if json_file is None:
        return tuple(no_change_15[:-1] + ["JSONファイルが選択されていません。"])

    try:
        with open(json_file.name, "r", encoding="utf-8") as f:
            settings = json.load(f)
    except Exception as e:
        return tuple(no_change_15[:-1] + [f"エラー: JSON読み込み失敗: {e}"])

    if not isinstance(settings, dict):
        return tuple(no_change_15[:-1] + ["エラー: 不正なJSON形式です。"])

    # 安全に取り出すユーティリティ
    def _get(key, default=None):
        return settings.get(key, default)

    # ドロップダウンの選択肢は現在のExcelに依存するため、
    # 取得した selected_targets は値のみ更新(choicesは触らない)
    selected = _get("selected_targets", []) or []
    spec_table = _get("spec_table", []) or []

    log = (f"設定をインポートしました(schema={_get('schema_version')}, "
           f"export_at={_get('exported_at')})。\n"
           f"※「解析対象」は現在のExcelの列/行ラベルと一致しないと反映されません。\n")

    return (
        gr.update(value=_get("subgroup_size", 5)),
        gr.update(value=_get("std_method", "サンプル標準偏差")),
        gr.update(value=_get("include_first_row", False)),
        gr.update(value=_get("include_first_column", False)),
        gr.update(value=_get("calc_direction", "列方向")),
        gr.update(value=_get("show_hist", True)),
        gr.update(value=_get("show_qq", True)),
        gr.update(value=_get("show_density", True)),
        gr.update(value=_get("show_xbar", True)),
        gr.update(value=_get("show_r", True)),
        gr.update(value=_get("show_s", True)),
        gr.update(value=_get("same_spec", False)),
        gr.update(value=selected),
        gr.update(value=spec_table),
        log
    )


# =========================================================================
# Gradio UI
# =========================================================================
with gr.Blocks(title="CpkTools-WebUI") as demo:
    gr.Markdown("# 🏭 CpkTools-WebUI 工程能力解析ツール")

    with gr.Tabs():
        # ============================
        # タブ1:基本統計量 (工程能力解析)
        # ============================
        with gr.Tab("📊 基本統計量"):
            gr.Markdown("### ① Excelファイルのアップロード")
            with gr.Row():
                file_input = gr.File(label="Excelファイル (xlsx, xls)", file_count="single")
            gr.Markdown("### ② データ読み込み設定")
            with gr.Row():
                include_first_row_chk = gr.Checkbox(
                    label="先頭行をデータとして扱う", value=False,
                    info="チェックすると、Excelの先頭行もデータとして読み込みます。"
                )
                include_first_column_chk = gr.Checkbox(
                    label="先頭列をデータとして扱う", value=False,
                    info="チェックすると、Excelの先頭列もデータとして読み込みます。"
                )
            with gr.Row():
                calc_direction_radio = gr.Radio(
                    choices=["列方向", "行方向"],
                    label="計算対象方向", value="列方向"
                )
            with gr.Row():
                preview_df = gr.DataFrame(
                    label="データプレビュー (先頭5行) - 先頭行/列の値はExcel記載のまま表示",
                    interactive=False
                )
            gr.Markdown("### ③ 解析対象の選択")
            with gr.Row():
                column_dropdown = gr.Dropdown(
                    choices=[], label="解析対象の列または行を選択", multiselect=True
                )
            gr.Markdown("### ④ 規格値の入力")
            gr.Markdown(
                "<small>⚠ 上限・下限の両方を空欄にすると、解析時にスキップされます。"
                "片側規格の場合は使用しない側のみ空欄にしてください。</small>"
            )
            with gr.Row():
                spec_df = gr.Dataframe(
                    headers=["解析対象", "規格上限値", "規格下限値"],
                    label="各対象の規格値入力(空欄は片側規格として扱います)",
                    interactive=True
                )
            with gr.Row():
                same_spec_chk = gr.Checkbox(
                    label="すべての対象の規格値を同じにする", value=False
                )
            gr.Markdown("### ⑤ 解析オプション")
            with gr.Row():
                subgroup_size_slider = gr.Slider(
                    minimum=1, maximum=10, step=1, value=5,
                    label="サブグループサイズ",
                    info="1の場合はI管理図/MR管理図、2以上ならX-bar/R/s管理図を生成"
                )
            with gr.Accordion("📈 出力グラフ設定", open=True):
                with gr.Row():
                    show_hist_checkbox = gr.Checkbox(label="ヒストグラムを出力", value=True)
                    show_qq_checkbox = gr.Checkbox(label="QQプロットを出力", value=True)
                    show_density_checkbox = gr.Checkbox(label="確率密度分布を出力", value=True)
                with gr.Row():
                    show_xbar_checkbox = gr.Checkbox(label="X-bar管理図／I管理図を出力", value=True)
                    show_r_checkbox = gr.Checkbox(label="R管理図／MR管理図を出力", value=True)
                    show_s_checkbox = gr.Checkbox(label="s管理図を出力", value=True)
            with gr.Row():
                std_method_radio = gr.Radio(
                    choices=["サンプル標準偏差", "母集団標準偏差"],
                    label="標準偏差の計算方法",
                    value="サンプル標準偏差"
                )

            # --- 設定のJSONインポート/エクスポート ---
            with gr.Accordion("⚙️ 設定のインポート / エクスポート (JSON)", open=False):
                with gr.Row():
                    settings_import_file = gr.File(
                        label="設定JSONをアップロードしてインポート",
                        file_types=[".json"], file_count="single"
                    )
                with gr.Row():
                    settings_import_button = gr.Button("📥 インポート実行", variant="secondary")
                    settings_export_button = gr.Button("📤 現在の設定をJSONエクスポート", variant="secondary")
                with gr.Row():
                    settings_export_file = gr.File(label="エクスポートされたJSONファイル")
                with gr.Row():
                    settings_log = gr.Textbox(
                        label="設定インポート/エクスポート ログ", lines=3, interactive=False
                    )

            gr.Markdown("### ⑥ 解析開始")
            run_button = gr.Button("解析開始", variant="primary")
            gr.Markdown("### 解析結果サマリー")
            cpk_result_html = gr.HTML()
            result_box = gr.Textbox(label="計算ログ", lines=8, interactive=False)
            with gr.Row():
                hist_gallery = gr.Gallery(label="ヒストグラム", show_label=True, height=300)
                qq_gallery = gr.Gallery(label="QQプロット", show_label=True, height=300)
            with gr.Row():
                density_gallery = gr.Gallery(label="確率密度分布", show_label=True, height=300)
            with gr.Row():
                xbar_gallery = gr.Gallery(label="X-bar管理図／I管理図", show_label=True, height=300)
                r_gallery = gr.Gallery(label="R管理図／MR管理図", show_label=True, height=300)
            with gr.Row():
                s_gallery = gr.Gallery(label="s管理図", show_label=True, height=300)
            with gr.Row():
                excel_file_box = gr.File(label="出力されたExcelファイル")
                excel_preview_box = gr.DataFrame(
                    label="Excelファイルの内容プレビュー", interactive=False
                )
            with gr.Row():
                open_folder_button = gr.Button("Outputフォルダを開く")

            # --- イベント接続 ---
            file_input.change(
                fn=update_preview,
                inputs=[file_input, include_first_row_chk, include_first_column_chk, calc_direction_radio],
                outputs=[preview_df, column_dropdown]
            )
            include_first_row_chk.change(
                fn=update_preview,
                inputs=[file_input, include_first_row_chk, include_first_column_chk, calc_direction_radio],
                outputs=[preview_df, column_dropdown]
            )
            include_first_column_chk.change(
                fn=update_preview,
                inputs=[file_input, include_first_row_chk, include_first_column_chk, calc_direction_radio],
                outputs=[preview_df, column_dropdown]
            )
            calc_direction_radio.change(
                fn=update_preview,
                inputs=[file_input, include_first_row_chk, include_first_column_chk, calc_direction_radio],
                outputs=[preview_df, column_dropdown]
            )
            column_dropdown.change(
                fn=update_spec_df_with_checkbox,
                inputs=[column_dropdown, same_spec_chk, spec_df],
                outputs=spec_df
            )
            same_spec_chk.change(
                fn=update_spec_df_with_checkbox,
                inputs=[column_dropdown, same_spec_chk, spec_df],
                outputs=spec_df
            )
            run_button.click(
                fn=run_analysis,
                inputs=[
                    file_input, column_dropdown, spec_df, subgroup_size_slider,
                    include_first_row_chk, include_first_column_chk, calc_direction_radio,
                    show_hist_checkbox, show_qq_checkbox, show_density_checkbox,
                    show_xbar_checkbox, show_r_checkbox, show_s_checkbox, std_method_radio
                ],
                outputs=[
                    result_box, cpk_result_html, hist_gallery, qq_gallery, density_gallery,
                    xbar_gallery, r_gallery, s_gallery, excel_file_box, excel_preview_box
                ]
            )
            open_folder_button.click(fn=open_output_folder, inputs=[], outputs=[])

            # --- 設定エクスポート ---
            settings_export_button.click(
                fn=export_settings_to_json,
                inputs=[
                    subgroup_size_slider, std_method_radio,
                    include_first_row_chk, include_first_column_chk, calc_direction_radio,
                    show_hist_checkbox, show_qq_checkbox, show_density_checkbox,
                    show_xbar_checkbox, show_r_checkbox, show_s_checkbox,
                    same_spec_chk, column_dropdown, spec_df
                ],
                outputs=[settings_export_file, settings_log]
            )

            # --- 設定インポート ---
            settings_import_button.click(
                fn=import_settings_from_json,
                inputs=[settings_import_file],
                outputs=[
                    subgroup_size_slider, std_method_radio,
                    include_first_row_chk, include_first_column_chk, calc_direction_radio,
                    show_hist_checkbox, show_qq_checkbox, show_density_checkbox,
                    show_xbar_checkbox, show_r_checkbox, show_s_checkbox,
                    same_spec_chk, column_dropdown, spec_df, settings_log
                ]
            )

        # ============================
        # タブ2:F検定 / T検定 / 相関
        # ============================
        with gr.Tab("🕵️ F検定/T検定/相関"):
            with gr.Row():
                test_file_input = gr.File(label="Excelファイル (xlsx, xls)", file_count="single")
            with gr.Row():
                include_first_row_chk_test = gr.Checkbox(
                    label="先頭行をデータとして扱う", value=False
                )
                include_first_column_chk_test = gr.Checkbox(
                    label="先頭列をデータとして扱う", value=False
                )
            with gr.Row():
                calc_direction_radio_test = gr.Radio(
                    choices=["列方向", "行方向"],
                    label="計算対象方向", value="列方向"
                )
            with gr.Row():
                preview_df_test = gr.DataFrame(
                    label="データプレビュー (先頭5行) - 先頭行/列の値はExcel記載のまま表示",
                    interactive=False
                )
            with gr.Row():
                test_column_dropdown = gr.Dropdown(
                    choices=[], label="検定対象の列または行を2つ選択", multiselect=True
                )
            with gr.Row():
                perform_f_test_radio = gr.Radio(
                    choices=["F検定を実施しない", "F検定を実施する"],
                    label="F検定", value="F検定を実施する"
                )
                alpha_f_input = gr.Number(label="有意水準 (F検定)", value=0.05, precision=3)
            with gr.Row():
                perform_t_test_radio = gr.Radio(
                    choices=["T検定を実施しない", "T検定を実施する"],
                    label="t検定の実行可否", value="T検定を実施する"
                )
                ttest_variant_radio = gr.Radio(
                    choices=["対応ありt検定", "独立t検定(分散が等しい)", "独立t検定(分散が異なる)"],
                    label="t検定の種類", value="対応ありt検定"
                )
                alpha_t_input = gr.Number(label="有意水準 (t検定)", value=0.05, precision=3)
            with gr.Row():
                plot_overlay_radio = gr.Radio(
                    choices=["正規分布を表示しない", "正規分布を表示する"],
                    label="正規分布の重ね描き",
                    value="正規分布を表示しない"
                )
                calc_corr_radio = gr.Radio(
                    choices=["相関計算を実施しない", "相関計算を実施する"],
                    label="相関計算",
                    value="相関計算を実施しない"
                )
                run_test_button = gr.Button("解析実行", variant="primary")
            with gr.Row():
                test_result_box = gr.Textbox(label="検定結果・ログ", lines=10, interactive=False)
            with gr.Row():
                density_overlay_gallery = gr.Gallery(label="理論分布プロット", show_label=True)
            with gr.Row():
                excel_test_file_box = gr.File(label="出力されたExcelファイル")
                excel_test_preview_box = gr.DataFrame(
                    label="Excelファイルの内容プレビュー", interactive=False
                )
            with gr.Row():
                open_folder_test_button = gr.Button("Outputフォルダを開く")

            test_file_input.change(
                fn=update_preview,
                inputs=[test_file_input, include_first_row_chk_test, include_first_column_chk_test, calc_direction_radio_test],
                outputs=[preview_df_test, test_column_dropdown]
            )
            include_first_row_chk_test.change(
                fn=update_preview,
                inputs=[test_file_input, include_first_row_chk_test, include_first_column_chk_test, calc_direction_radio_test],
                outputs=[preview_df_test, test_column_dropdown]
            )
            include_first_column_chk_test.change(
                fn=update_preview,
                inputs=[test_file_input, include_first_row_chk_test, include_first_column_chk_test, calc_direction_radio_test],
                outputs=[preview_df_test, test_column_dropdown]
            )
            calc_direction_radio_test.change(
                fn=update_preview,
                inputs=[test_file_input, include_first_row_chk_test, include_first_column_chk_test, calc_direction_radio_test],
                outputs=[preview_df_test, test_column_dropdown]
            )

            run_test_button.click(
                fn=run_stat_test,
                inputs=[
                    test_file_input, test_column_dropdown,
                    perform_f_test_radio, alpha_f_input,
                    perform_t_test_radio, ttest_variant_radio, alpha_t_input,
                    include_first_row_chk_test, include_first_column_chk_test,
                    plot_overlay_radio, calc_corr_radio, calc_direction_radio_test
                ],
                outputs=[test_result_box, density_overlay_gallery, excel_test_file_box, excel_test_preview_box]
            )
            open_folder_test_button.click(fn=open_output_folder, inputs=[], outputs=[])

        # ============================
        # タブ3: 初学者向け解説
        # ============================
        with gr.Tab("📖 初学者向け解説"):
            try:
                with open("explanation.txt", "r", encoding="utf-8") as f:
                    explanation_text = f.read()
            except Exception as e:
                explanation_text = f"解説ファイルの読み込みに失敗しました: {e}"
            gr.Markdown(explanation_text)

    # フッター - GitHub リンクとバージョン情報
    version = get_version()
    gr.Markdown(
        f"---\n"
        f"©2025 @KotaOoka  |  **バージョン: {version}**  |  "
        f"GitHub: [{GITHUB_URL}]({GITHUB_URL})"
    )

if __name__ == "__main__":
    demo.launch(inbrowser=True, share=False)

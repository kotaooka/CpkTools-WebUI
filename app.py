import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt
import datetime
import os
import numpy as np
import math
from PIL import Image
import gradio as gr
import re

# -------------------------
# ファイル名サニタイズ関数
# ※ひらがな: \u3040-\u309F, カタカナ: \u30A0-\u30FF, 漢字: \u4E00-\u9FFF, 全角数字: \uFF10-\uFF19 も許可
def sanitize_filename(name):
    allowed_pattern = r'[^A-Za-z0-9\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF\uFF10-\uFF19_\-]'
    return re.sub(allowed_pattern, '_', name, flags=re.UNICODE)

# -------------------------
# 初期設定：日本語フォント設定（Windowsの場合）
plt.rcParams['font.family'] = 'MS Gothic'

# -------------------------
# 出力先ディレクトリの設定（固定）
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def get_version():
    try:
        with open("version.txt", "r") as f:
            return f.read().strip()
    except Exception:
        return "バージョン情報未設定"

# -------------------------
# Excelファイルの読み込み共通関数
def read_excel_file(uploaded_file, include_first_row, include_first_column):
    """
    include_first_row が True の場合はすべての行をデータとして扱い、
    False の場合は先頭行をヘッダーとして読み込む。
    include_first_column が True の場合はすべての列をデータとして扱い、
    False の場合は先頭列をインデックスとして読み込み、データとしては表示しない。
    """
    header = None if include_first_row else 0
    index_col = None if include_first_column else 0
    df = pd.read_excel(uploaded_file.name, header=header, index_col=index_col)
    return df

# -------------------------
# 補助関数：選択された列（または行）に合わせて規格値テーブルを更新する（インタラクティブ対応）
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
        if idx < len(current_spec_list) and current_spec_list[idx] and current_spec_list[idx][0] == target:
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

    # プレビュー表示用に列名を常に「列A, 列B, …」に上書きする
    df.columns = [f"列{chr(65+i)}" for i in range(len(df.columns))]
    if include_first_column:
        df.index = [f"行{j+1}" for j in range(len(df))]
    else:
        # 先頭列がデータとして扱われない場合は、reset_index()してインデックスを除去
        df = df.reset_index(drop=True)
        df.index = [f"行{j+1}" for j in range(len(df))]
    
    # プレビューは先頭5行のみ表示（表示上は制限するが、選択肢は全体）
    preview_df = df.head(5)
    
    # 解析対象選択肢は、計算対象方向に合わせ全体の列（または行）のラベルを返す
    if calc_direction == "列方向":
        target_choices = list(df.columns)
    else:
        target_choices = list(df.index)
    
    return preview_df, gr.update(choices=target_choices)

# -------------------------
# 解析処理（工程能力解析ツール）
def run_analysis(uploaded_file, selected_targets, spec_table, subgroup_size, include_first_row, include_first_column,
                 calc_direction, show_hist, show_qq, show_density, show_xbar, show_r, show_s, std_method):
    log_messages = ""
    hist_images = []      # ヒストグラム
    qq_images = []        # QQプロット
    density_images = []   # 確率密度分布
    xbar_images = []      # X-bar管理図 (または I管理図)
    r_images = []         # R管理図 (または MR管理図)
    s_images = []         # s管理図（サブグループサイズ>=2の場合のみ）
    excel_file = None     # Excel出力ファイルパス
    excel_preview = None  # Excel出力結果プレビュー
    results = []          # 各対象の統計解析結果リスト

    ddof_value = 1 if std_method == "サンプル標準偏差" else 0

    if uploaded_file is None:
        return "エラー: ファイルが選択されていません", None, None, None, None, None, None, None, None
    try:
        df = read_excel_file(uploaded_file, include_first_row, include_first_column)
        log_messages += "ファイル読み込み成功。\n"
    except Exception as e:
        return f"エラー: ファイル読み込み中に問題が発生しました: {e}", None, None, None, None, None, None, None, None

    if df.empty:
        return "エラー: ファイルにデータがありません", None, None, None, None, None, None, None, None
    if not selected_targets:
        return "エラー: 解析対象が選択されていません", None, None, None, None, None, None, None, None
    try:
        spec_df = pd.DataFrame(spec_table, columns=["解析対象", "規格上限値", "規格下限値"])
    except Exception as e:
        return f"エラー: 規格値テーブルの読み込みに失敗しました: {e}", None, None, None, None, None, None, None, None

    if len(spec_df) != len(selected_targets):
        return "エラー: 選択された対象数と規格値入力の行数が一致しません。", None, None, None, None, None, None, None, None

    for i, target_label in enumerate(selected_targets):
        # 対象データの抽出（計算対象方向に応じ、列または行を選択）
        if calc_direction == "列方向":
            try:
                column_index = ord(target_label.replace("列", "")) - 65
            except Exception as e:
                log_messages += f"エラー: 選択された列の形式が正しくありません ({target_label})\n"
                continue
            if column_index < 0 or column_index >= len(df.columns):
                log_messages += f"エラー: 正しい列を選択してください ({target_label})\n"
                continue
            actual_label = df.columns[column_index]
            data = df[actual_label]
        else:
            try:
                m = re.match(r"行\s*(\d+)", target_label)
                if m:
                    row_index = int(m.group(1)) - 1
                else:
                    raise ValueError("行番号の解析失敗")
            except Exception as e:
                log_messages += f"エラー: 選択された行の形式が正しくありません ({target_label})\n"
                continue
            if row_index < 0 or row_index >= len(df):
                log_messages += f"エラー: 正しい行を選択してください ({target_label})\n"
                continue
            data = df.iloc[row_index]
            actual_label = f"Row {row_index+1}"

        # 欠損値の除去
        if data.isnull().any():
            log_messages += f"注意: {target_label} に欠損値が存在します。欠損値を除外して解析します。（元のデータ数: {len(data)}）\n"
            data = data.dropna()
        sample_n = len(data)
        if sample_n == 0:
            log_messages += f"エラー: {target_label} のデータがすべて欠損しています。\n"
            continue
        try:
            spec_entry_usl = spec_df.iloc[i]["規格上限値"]
            spec_entry_lsl = spec_df.iloc[i]["規格下限値"]
            current_usl = float(spec_entry_usl) if str(spec_entry_usl).strip() != "" else None
            current_lsl = float(spec_entry_lsl) if str(spec_entry_lsl).strip() != "" else None
        except Exception as e:
            log_messages += f"エラー: {target_label} の規格値が正しく入力されていません: {e}\n"
            continue
        try:
            max_val = float(data.max())
            min_val = float(data.min())
            std_val = float(data.std(ddof=ddof_value))
            mean_val = float(data.mean())
            kurtosis_val = float(data.kurtosis())
            skewness_val = float(data.skew())
            if std_val == 0:
                log_messages += f"エラー: {target_label} の標準偏差が0のため、Cp/Cpk計算をスキップしました。（サンプル数: {sample_n}）\n"
                continue
            if current_usl is not None and current_lsl is not None:
                spec_type = "両側"
                Cp = (current_usl - current_lsl) / (6 * std_val)
                Cpk = min((current_usl - mean_val), (mean_val - current_lsl)) / (3 * std_val)
            elif current_usl is not None:
                spec_type = "上側のみ"
                Cp = (current_usl - mean_val) / (3 * std_val)
                Cpk = (current_usl - mean_val) / (3 * std_val)
            elif current_lsl is not None:
                spec_type = "下側のみ"
                Cp = (mean_val - current_lsl) / (3 * std_val)
                Cpk = (mean_val - current_lsl) / (3 * std_val)
            else:
                log_messages += f"エラー: {target_label} の規格値が入力されていません。\n"
                continue

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
                "Cpk": Cpk,
                "尖度": kurtosis_val,
                "歪度": skewness_val
            })
            log_messages += f"解析対象: {target_label} ({actual_label}) の統計計算完了。（サンプル数: {sample_n}）\n"
        except Exception as e:
            log_messages += f"エラー: {target_label} の統計計算中に問題が発生しました: {e}\n"
            continue

        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_col_label = sanitize_filename(target_label)

        # ヒストグラム生成
        if show_hist:
            try:
                plt.figure()
                plt.hist(data, color="skyblue", edgecolor="black")
                plt.xlabel("値")
                plt.ylabel("度数")
                plt.title(f"ヒストグラム ({target_label})")
                hist_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_hist_{target_label}.jpg")
                plt.savefig(hist_filename, format="jpg")
                plt.close()
                hist_images.append(hist_filename)
                log_messages += f"{target_label} のヒストグラム生成完了。\n"
            except Exception as e:
                log_messages += f"エラー: {target_label} のヒストグラム生成中に問題が発生しました: {e}\n"

        # QQプロット生成
        if show_qq:
            try:
                plt.figure()
                stats.probplot(data, dist="norm", plot=plt)
                plt.title(f"QQプロット ({target_label})")
                qq_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_qq_{target_label}.jpg")
                plt.savefig(qq_filename, format="jpg")
                plt.close()
                qq_images.append(qq_filename)
                log_messages += f"{target_label} のQQプロット生成完了。\n"
            except Exception as e:
                log_messages += f"エラー: {target_label} のQQプロット生成中に問題が発生しました: {e}\n"

        # 確率密度分布生成（正規分布フィッティング）
        if show_density:
            try:
                plt.figure()
                x_vals = np.linspace(mean_val - 4 * std_val, mean_val + 4 * std_val, 100)
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
                ax.text(mean_val - 3 * std_val, label_y, f"-3σ: {mean_val - 3 * std_val:.2f}", rotation=90,
                        color="black", ha="center", va="bottom", fontsize=8)
                ax.text(mean_val + 3 * std_val, label_y, f"+3σ: {mean_val + 3 * std_val:.2f}", rotation=90,
                        color="black", ha="center", va="bottom", fontsize=8)
                ax.text(mean_val, label_y, f"平均値: {mean_val:.2f}", rotation=90,
                        color="black", ha="center", va="bottom", fontsize=8)
                if current_usl is not None:
                    ax.text(current_usl, label_y, f"規格上限値: {current_usl:.2f}", rotation=90,
                            color="black", ha="center", va="bottom", fontsize=8)
                if current_lsl is not None:
                    ax.text(current_lsl, label_y, f"規格下限値: {current_lsl:.2f}", rotation=90,
                            color="black", ha="center", va="bottom", fontsize=8)
                plt.xlabel("値")
                plt.ylabel("確率密度")
                plt.title(f"確率密度分布 ({target_label})")
                plt.legend()
                density_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_density_{target_label}.jpg")
                plt.savefig(density_filename, format="jpg")
                plt.close()
                density_images.append(density_filename)
                log_messages += f"{target_label} の確率密度分布描画完了。\n"
            except Exception as e:
                log_messages += f"エラー: {target_label} の確率密度分布描画中に問題が発生しました: {e}\n"

        # サブグループ管理図（I管理図/MR管理図、またはX-bar/R/s管理図）
        if subgroup_size == 1:
            individuals = data.values
            n_individuals = len(individuals)
            if n_individuals < 1:
                log_messages += f"警告: {target_label} のデータ点数が不足しているため、管理図を生成できませんでした。\n"
            else:
                i_bar = np.mean(individuals)
                if n_individuals >= 2:
                    moving_ranges = [abs(individuals[j] - individuals[j-1]) for j in range(1, n_individuals)]
                    mr_bar = np.mean(moving_ranges)
                    sigma = mr_bar / 1.128
                else:
                    moving_ranges = []
                    mr_bar = 0
                    sigma = 0
                UCL_i = i_bar + 3 * sigma
                LCL_i = i_bar - 3 * sigma
                if show_xbar:
                    try:
                        plt.figure()
                        plt.plot(range(1, n_individuals+1), individuals, marker='o', linestyle='-', color='blue', label='値')
                        plt.axhline(i_bar, color='green', linestyle='--', label='平均')
                        plt.axhline(UCL_i, color='red', linestyle='--', label='UCL')
                        plt.axhline(LCL_i, color='red', linestyle='--', label='LCL')
                        plt.xlabel('データ点')
                        plt.ylabel('値')
                        plt.title(f"I管理図 ({target_label})")
                        plt.legend()
                        i_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_i_{target_label}.jpg")
                        plt.savefig(i_filename, format="jpg")
                        plt.close()
                        xbar_images.append(i_filename)
                        log_messages += f"{target_label} のI管理図生成完了。\n"
                    except Exception as e:
                        log_messages += f"エラー: {target_label} のI管理図生成中に問題が発生しました: {e}\n"
                if show_r:
                    if n_individuals >= 2:
                        try:
                            plt.figure()
                            plt.plot(range(2, n_individuals+1), moving_ranges, marker='o', linestyle='-', color='blue', label='移動範囲')
                            plt.axhline(mr_bar, color='green', linestyle='--', label='平均MR')
                            UCL_mr = mr_bar * 3.267
                            plt.axhline(UCL_mr, color='red', linestyle='--', label='UCL')
                            plt.xlabel('データ点 (2番目以降)')
                            plt.ylabel('移動範囲')
                            plt.title(f"MR管理図 ({target_label})")
                            plt.legend()
                            mr_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_mr_{target_label}.jpg")
                            plt.savefig(mr_filename, format="jpg")
                            plt.close()
                            r_images.append(mr_filename)
                            log_messages += f"{target_label} のMR管理図生成完了。\n"
                        except Exception as e:
                            log_messages += f"エラー: {target_label} のMR管理図生成中に問題が発生しました: {e}\n"
                    else:
                        log_messages += f"警告: {target_label} のデータ点数が不十分なため、MR管理図を生成できませんでした。\n"
                if show_s:
                    log_messages += f"警告: サブグループサイズが1のため、s管理図は生成できません。\n"
        else:
            # サブグループサイズが1以外の場合：X-bar管理図、R管理図、s管理図の生成
            groups = [data[i:i+subgroup_size] for i in range(0, len(data), subgroup_size)]
            full_groups = [group for group in groups if len(group) == subgroup_size]
            if not full_groups:
                log_messages += f"警告: {target_label} のサブグループが形成できませんでした。\n"
            else:
                group_means = [np.mean(group) for group in full_groups]
                group_ranges = [np.max(group) - np.min(group) for group in full_groups]
                group_std = [np.std(group, ddof=1) for group in full_groups]
                Xbar_center = np.mean(group_means)
                Rbar = np.mean(group_ranges)
                sbar = np.mean(group_std)
                
                subgroup_n = subgroup_size
                factors = {
                    2: {'A2': 1.88, 'D3': 0, 'D4': 3.267, 'B3': 0, 'B4': 2.568},
                    3: {'A2': 1.023, 'D3': 0, 'D4': 2.574, 'B3': 0, 'B4': 2.089},
                    4: {'A2': 0.729, 'D3': 0, 'D4': 2.282, 'B3': 0, 'B4': 1.880},
                    5: {'A2': 0.577, 'D3': 0, 'D4': 2.114, 'B3': 0, 'B4': 1.716},
                    6: {'A2': 0.483, 'D3': 0, 'D4': 2.004, 'B3': 0.03, 'B4': 1.924},
                    7: {'A2': 0.419, 'D3': 0.076, 'D4': 1.924, 'B3': 0.118, 'B4': 1.816},
                    8: {'A2': 0.373, 'D3': 0.136, 'D4': 1.864, 'B3': 0.185, 'B4': 1.734},
                    9: {'A2': 0.337, 'D3': 0.184, 'D4': 1.816, 'B3': 0.239, 'B4': 1.66},
                    10:{'A2': 0.308, 'D3': 0.223, 'D4': 1.777, 'B3': 0.284, 'B4': 1.602},
                }
                if subgroup_n not in factors:
                    subgroup_n = max(factors.keys())
                current_factors = factors[subgroup_n]
                
                # X-bar Chart
                if show_xbar:
                    try:
                        UCL_xbar = Xbar_center + current_factors['A2'] * Rbar
                        LCL_xbar = Xbar_center - current_factors['A2'] * Rbar
                        plt.figure()
                        plt.plot(range(1, len(group_means)+1), group_means, marker='o', linestyle='-', color='blue', label='サブグループ平均')
                        plt.axhline(Xbar_center, color='green', linestyle='--', label='全体平均')
                        plt.axhline(UCL_xbar, color='red', linestyle='--', label='UCL')
                        plt.axhline(LCL_xbar, color='red', linestyle='--', label='LCL')
                        plt.xlabel('サブグループ番号')
                        plt.ylabel('サブグループ平均')
                        plt.title(f"X-bar管理図 ({target_label})")
                        plt.legend()
                        xbar_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_xbar_{sanitize_filename(target_label)}.jpg")
                        plt.savefig(xbar_filename, format="jpg")
                        plt.close()
                        xbar_images.append(xbar_filename)
                        log_messages += f"{target_label} のX-bar管理図生成完了。\n"
                    except Exception as e:
                        log_messages += f"エラー: {target_label} のX-bar管理図生成中に問題が発生しました: {e}\n"
                # R Chart
                if show_r:
                    try:
                        UCL_R = current_factors['D4'] * Rbar
                        LCL_R = current_factors['D3'] * Rbar
                        plt.figure()
                        plt.plot(range(1, len(group_ranges)+1), group_ranges, marker='o', linestyle='-', color='blue', label='サブグループレンジ')
                        plt.axhline(Rbar, color='green', linestyle='--', label='平均レンジ')
                        plt.axhline(UCL_R, color='red', linestyle='--', label='UCL')
                        plt.axhline(LCL_R, color='red', linestyle='--', label='LCL')
                        plt.xlabel('サブグループ番号')
                        plt.ylabel('レンジ')
                        plt.title(f"R管理図 ({target_label})")
                        plt.legend()
                        r_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_r_{sanitize_filename(target_label)}.jpg")
                        plt.savefig(r_filename, format="jpg")
                        plt.close()
                        r_images.append(r_filename)
                        log_messages += f"{target_label} のR管理図生成完了。\n"
                    except Exception as e:
                        log_messages += f"エラー: {target_label} のR管理図生成中に問題が発生しました: {e}\n"
                # s Chart
                if show_s:
                    try:
                        UCL_s = sbar * current_factors['B4']
                        LCL_s = sbar * current_factors['B3']
                        plt.figure()
                        plt.plot(range(1, len(group_std)+1), group_std, marker='o', linestyle='-', color='blue', label='サブグループ標準偏差')
                        plt.axhline(sbar, color='green', linestyle='--', label='平均標準偏差')
                        plt.axhline(UCL_s, color='red', linestyle='--', label='UCL')
                        plt.axhline(LCL_s, color='red', linestyle='--', label='LCL')
                        plt.xlabel('サブグループ番号')
                        plt.ylabel('標準偏差')
                        plt.title(f"s管理図 ({target_label})")
                        plt.legend()
                        s_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_s_{sanitize_filename(target_label)}.jpg")
                        plt.savefig(s_filename, format="jpg")
                        plt.close()
                        s_images.append(s_filename)
                        log_messages += f"{target_label} のs管理図生成完了。\n"
                    except Exception as e:
                        log_messages += f"エラー: {target_label} のs管理図生成中に問題が発生しました: {e}\n"

    if results:
        dt_now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = os.path.join(OUTPUT_DIR, f"{dt_now}_results.xlsx")
        try:
            results_df = pd.DataFrame(results)
            results_df.to_excel(output_filename, index=False)
            excel_file = output_filename
            excel_preview = results_df
            log_messages += f"Excelファイル出力完了: {output_filename}\n"
        except Exception as e:
            log_messages += f"エラー: Excelファイル書き出し中に問題が発生しました: {e}\n"
    else:
        log_messages += "エラー: 解析対象から有効なデータが得られませんでした。\n"

    return log_messages, hist_images, qq_images, density_images, xbar_images, r_images, s_images, excel_file, excel_preview

# -------------------------
# F検定/t検定実施関数（テストタブ用）
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

    if uploaded_file is None:
        return "エラー: ファイルが選択されていません", [], None, None

    try:
        df = read_excel_file(uploaded_file, include_first_row, include_first_column)
        log_messages += "ファイル読み込み成功。\n"
    except Exception as e:
        return f"エラー: ファイル読み込み中に問題が発生しました: {e}", [], None, None

    if df.empty:
        return "エラー: ファイルにデータがありません", [], None, None

    if len(selected_targets) != 2:
        return "エラー: 検定対象は2つ選択してください", [], None, None

    data_list = []
    col_names = []
    for target_label in selected_targets:
        if calc_direction == "列方向":
            try:
                column_index = ord(target_label.replace("列", "")) - 65
            except Exception as e:
                log_messages += f"エラー: 選択された列の形式が正しくありません: {target_label}\n"
                continue
            if column_index < 0 or column_index >= len(df.columns):
                log_messages += f"エラー: 正しい列を選択してください: {target_label}\n"
                continue
            actual_label = df.columns[column_index]
            col_names.append(target_label)
            col_data = df[actual_label].dropna()
        else:
            try:
                m = re.match(r"行\s*(\d+)", target_label)
                if m:
                    row_index = int(m.group(1)) - 1
                else:
                    raise ValueError("行番号の解析失敗")
            except Exception as e:
                log_messages += f"エラー: 選択された行の形式が正しくありません: {target_label}\n"
                continue
            if row_index < 0 or row_index >= len(df):
                log_messages += f"エラー: 正しい行を選択してください: {target_label}\n"
                continue
            col_names.append(target_label)
            col_data = df.iloc[row_index].dropna()
        if len(col_data) == 0:
            log_messages += f"エラー: {target_label} のデータが全て欠損しています。\n"
        data_list.append(col_data)
    if len(data_list) != 2:
        return "エラー: 選択された2対象のデータに問題があります。", [], None, None

    data1, data2 = data_list[0], data_list[1]
    n1, n2 = len(data1), len(data2)
    mean1, mean2 = data1.mean(), data2.mean()
    var1, var2 = data1.var(ddof=1), data2.var(ddof=1)

    # ------ F検定 ------
    if perform_f_test == "F検定を実施する":
        if var1 >= var2:
            f_stat = var1 / var2
            dfn = n1 - 1
            dfd = n2 - 1
        else:
            f_stat = var2 / var1
            dfn = n2 - 1
            dfd = n1 - 1
        p_value_f = 2 * min(stats.f.cdf(f_stat, dfn, dfd), 1 - stats.f.cdf(f_stat, dfn, dfd))
        significance_f = "有意差あり" if p_value_f < alpha_f else "有意差なし"
        log_messages += (
            f"F検定結果:\n"
            f"サンプル数: {n1}, {n2}\n"
            f"平均: {mean1:.4f}, {mean2:.4f}\n"
            f"分散: {var1:.4f}, {var2:.4f}\n"
            f"F値: {f_stat:.4f}\n"
            f"自由度: ({dfn}, {dfd})\n"
            f"P値: {p_value_f:.4f}\n"
            f"判定: {significance_f}\n"
        )
        f_test_done = True

    # ------ t検定 ------
    if perform_t_test != "T検定を実施しない":
        if ttest_variant == "対応ありt検定":
            df_pair = pd.DataFrame({"data1": data1, "data2": data2}).dropna()
            if df_pair.empty:
                return "エラー: 両群で有効なペアが存在しません。", [], None, None
            data1 = df_pair["data1"]
            data2 = df_pair["data2"]
            n1 = len(data1)
            t_stat, p_value_t = stats.ttest_rel(data1, data2)
            df_t = n1 - 1
        elif ttest_variant == "独立t検定（分散が等しい）":
            t_stat, p_value_t = stats.ttest_ind(data1, data2, equal_var=True)
            df_t = n1 + n2 - 2
        elif ttest_variant == "独立t検定（分散が異なる）":
            t_stat, p_value_t = stats.ttest_ind(data1, data2, equal_var=False)
            df_t = ((var1/n1 + var2/n2)**2 / ((var1/n1)**2/(n1-1) + (var2/n2)**2/(n2-1)))
        else:
            return "エラー: t検定の種類が選択されていません。", [], None, None
        significance_t = "有意差あり" if p_value_t < alpha_t else "有意差なし"
        log_messages += (
            f"t検定結果 ({ttest_variant}):\n"
            f"サンプル数: {n1}, {n2}\n"
            f"平均: {mean1:.4f}, {mean2:.4f}\n"
            f"分散: {var1:.4f}, {var2:.4f}\n"
            f"t値: {t_stat:.4f}\n"
            f"P値: {p_value_t:.4f}\n"
            f"有意水準: {alpha_t}\n"
            f"判定: {significance_t}\n"
        )
        t_test_done = True

    result_dict = {"検定対象1": col_names[0], "検定対象2": col_names[1],"サンプル数1": n1,"サンプル数2": n2,"平均1": mean1,"平均2": mean2,"分散1": var1,"分散2": var2}



    if f_test_done:
        result_dict.update({
            "F値": f_stat,
            "P値 (F検定)": p_value_f,
            "判定 (F検定)": significance_f
        })
    if t_test_done:
        result_dict.update({
            "t値": t_stat,
            "P値 (t検定)": p_value_t,
            "判定 (t検定)": significance_t
        })

    # --- 相関計算 ---
    if calc_corr == "相関計算を実施する":
        corr_coeff = None
        r2_value = None
        try:
            df_corr = pd.DataFrame({"data1": data1, "data2": data2}).dropna()
            if len(df_corr) > 1:
                corr_coeff = df_corr["data1"].corr(df_corr["data2"])
                r2_value = corr_coeff ** 2
                log_messages += f"相関係数: {corr_coeff:.4f}\n決定係数: {r2_value:.4f}\n"
            else:
                log_messages += "警告: 相関計算に十分なデータがありません。\n"
        except Exception as e:
            log_messages += f"エラー: 相関計算中に問題が発生しました: {e}\n"

        if corr_coeff is not None:
            result_dict.update({"相関係数": corr_coeff, "決定係数": r2_value})

        try:
            plt.figure()
            plt.scatter(df_corr["data1"], df_corr["data2"], color="blue", label="データポイント")
            slope, intercept = np.polyfit(df_corr["data1"], df_corr["data2"], 1)
            x_vals = np.linspace(df_corr["data1"].min(), df_corr["data1"].max(), 100)
            y_vals = slope * x_vals + intercept
            plt.plot(x_vals, y_vals, color="red", label="回帰直線")
            plt.xlabel(selected_targets[0])
            plt.ylabel(selected_targets[1])
            plt.title("散布図")
            plt.legend()
            dt_scatter = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            scatter_filename = os.path.join(OUTPUT_DIR, f"{dt_scatter}_scatter.jpg")
            plt.savefig(scatter_filename, format="jpg")
            plt.close()
            density_images.append(scatter_filename)
            log_messages += "散布図生成完了。\n"
        except Exception as e:
            log_messages += f"エラー: 散布図生成中に問題が発生しました: {e}\n"

    results_df = pd.DataFrame([result_dict])
    dt_now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_filename = os.path.join(OUTPUT_DIR, f"{dt_now}_stat_test_results.xlsx")
    try:
        results_df.to_excel(excel_filename, index=False)
        excel_file = excel_filename
        excel_preview = results_df
    except Exception as e:
        log_messages += f"エラー: Excelファイル書き出し中に問題が発生しました: {e}\n"

    # ------ グラフ生成（正規分布重ね描きなど） ------
    if plot_overlay == "正規分布を表示する":
        dt_now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        try:
            overall_min = min(data1.min(), data2.min())
            overall_max = max(data1.max(), data2.max())
            range_padding = (overall_max - overall_min) * 0.1
            x_vals = np.linspace(overall_min - range_padding, overall_max + range_padding, 200)
            std1 = data1.std(ddof=1)
            std2 = data2.std(ddof=1)
            y1_vals = stats.norm.pdf(x_vals, loc=mean1, scale=std1)
            y2_vals = stats.norm.pdf(x_vals, loc=mean2, scale=std2)
            plt.figure()
            plt.plot(x_vals, y1_vals, label=f"{selected_targets[0]}", color="blue")
            plt.plot(x_vals, y2_vals, label=f"{selected_targets[1]}", color="red")
            plt.xlabel("値")
            plt.ylabel("確率密度")
            plt.title("各群正規分布の重ね描き")
            plt.legend()
            overlay_filename = os.path.join(OUTPUT_DIR, f"{dt_now}_density_overlay.jpg")
            plt.savefig(overlay_filename, format="jpg")
            plt.close()
            density_images.append(overlay_filename)
            log_messages += "各群正規分布の重ね描きプロット生成完了。\n"
        except Exception as e:
            log_messages += f"エラー: 各群正規分布の重ね描きプロット生成中に問題が発生しました: {e}\n"

    if t_test_done:
        try:
            lower_bound = stats.t.ppf(0.001, df_t)
            upper_bound = stats.t.ppf(0.999, df_t)
            x_t = np.linspace(lower_bound, upper_bound, 200)
            y_t = stats.t.pdf(x_t, df_t)
            plt.figure()
            plt.plot(x_t, y_t, label=f"t分布 (df={df_t:.2f})", color="purple")
            plt.axvline(t_stat, color="black", linestyle="--", label=f"t値 = {t_stat:.2f}")
            plt.text(t_stat, max(y_t)*0.7, f"p={p_value_t:.3f}", color="black", fontsize=10,
                     rotation=90, ha="left", va="center")
            plt.xlabel("t値")
            plt.ylabel("確率密度")
            plt.title("t分布プロット")
            plt.legend()
            timestamp2 = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            t_plot_filename = os.path.join(OUTPUT_DIR, f"{timestamp2}_t_distribution.jpg")
            plt.savefig(t_plot_filename, format="jpg")
            plt.close()
            density_images.append(t_plot_filename)
            log_messages += "t分布プロット生成完了。\n"
        except Exception as e:
            log_messages += f"エラー: t分布プロット生成中に問題が発生しました: {e}\n"
    if f_test_done:
        try:
            lower_bound = stats.f.ppf(0.001, dfn, dfd)
            upper_bound = stats.f.ppf(0.999, dfn, dfd)
            x_f = np.linspace(lower_bound, upper_bound, 200)
            y_f = stats.f.pdf(x_f, dfn, dfd)
            plt.figure()
            plt.plot(x_f, y_f, label=f"F分布 (dfn={dfn}, dfd={dfd})", color="orange")
            plt.axvline(f_stat, color="black", linestyle="--", label=f"F値 = {f_stat:.2f}")
            plt.text(f_stat, max(y_f)*0.7, f"p={p_value_f:.3f}", color="black", fontsize=10,
                     rotation=90, ha="left", va="center")
            plt.xlabel("F値")
            plt.ylabel("確率密度")
            plt.title("F分布プロット")
            plt.legend()
            timestamp3 = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            f_plot_filename = os.path.join(OUTPUT_DIR, f"{timestamp3}_f_distribution.jpg")
            plt.savefig(f_plot_filename, format="jpg")
            plt.close()
            density_images.append(f_plot_filename)
            log_messages += "F分布プロット生成完了。\n"
        except Exception as e:
            log_messages += f"エラー: F分布プロット生成中に問題が発生しました: {e}\n"

    return log_messages, density_images, excel_file, excel_preview

# -------------------------
# Outputフォルダを開く関数（Windows専用）
def open_output_folder():
    folder_path = os.path.abspath(OUTPUT_DIR)
    try:
        os.startfile(folder_path)
    except Exception as e:
        print(f"フォルダを開くのに失敗しました: {e}")

# -------------------------
# Gradio UI の構築
with gr.Blocks() as demo:
    gr.Markdown("# 🏭 CpkTools-WebUI 工程能力解析ツール")

    with gr.Tabs():
        # タブ1：基本統計量（工程能力解析）
        with gr.Tab("📊 基本統計量"):
            with gr.Row():
                file_input = gr.File(label="Excelファイル (xlsx, xls)", file_count="single")
            with gr.Row():
                include_first_row_chk = gr.Checkbox(
                    label="先頭行をデータとして扱う", value=False,
                    info="チェックすると、Excelの先頭行もデータとして読み込みます。"
                )
            with gr.Row():
                include_first_column_chk = gr.Checkbox(
                    label="先頭列をデータとして扱う", value=False,
                    info="チェックすると、Excelの先頭列もデータとして読み込みます。"
                )
            with gr.Row():
                calc_direction_radio = gr.Radio(choices=["列方向", "行方向"],
                                                  label="計算対象方向", value="列方向")
            with gr.Row():
                preview_df = gr.DataFrame(label="データプレビュー (先頭5行)", interactive=False)
            with gr.Row():
                column_dropdown = gr.Dropdown(choices=[], label="解析対象の列または行を選択", multiselect=True)
            with gr.Row():
                spec_df = gr.Dataframe(
                    headers=["解析対象", "規格上限値", "規格下限値"],
                    label="各対象の規格値入力（空欄は片側規格として扱います）", interactive=True
                )
            with gr.Row():
                same_spec_chk = gr.Checkbox(label="すべての対象の規格値を同じにする", value=False)
            with gr.Row():
                subgroup_size_slider = gr.Slider(
                    minimum=1, maximum=10, step=1, value=5,
                    label="サブグループサイズ",
                    info="X-bar管理図、R管理図、s管理図作成時に用いるサブグループのサイズ。1の場合はI管理図/MR管理図を生成します。"
                )
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
            run_button = gr.Button("解析開始")
            result_box = gr.Textbox(label="計算結果・ログ", lines=10, interactive=False)
            with gr.Row():
                hist_gallery = gr.Gallery(label="ヒストグラム", show_label=True, type="file")
                qq_gallery = gr.Gallery(label="QQプロット", show_label=True, type="file")
            with gr.Row():
                density_gallery = gr.Gallery(label="確率密度分布", show_label=True, type="file")
            with gr.Row():
                xbar_gallery = gr.Gallery(label="X-bar管理図／I管理図", show_label=True, type="file")
                r_gallery = gr.Gallery(label="R管理図／MR管理図", show_label=True, type="file")
            with gr.Row():
                s_gallery = gr.Gallery(label="s管理図", show_label=True, type="file")
            with gr.Row():
                excel_file_box = gr.File(label="出力されたExcelファイルを開く")
                excel_preview_box = gr.DataFrame(label="Excelファイルの内容プレビュー", interactive=False)
            with gr.Row():
                open_folder_button = gr.Button("Outputフォルダを開く")

            file_input.change(
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
                    result_box, hist_gallery, qq_gallery, density_gallery,
                    xbar_gallery, r_gallery, s_gallery, excel_file_box, excel_preview_box
                ]
            )
            open_folder_button.click(fn=open_output_folder, inputs=[], outputs=[])

        # タブ2：F検定/T検定/相関
        with gr.Tab("🕵️F検定/T検定/相関"):
            with gr.Row():
                test_file_input = gr.File(label="Excelファイル (xlsx, xls)", file_count="single")
            with gr.Row():
                include_first_row_chk_test = gr.Checkbox(
                    label="先頭行をデータとして扱う", value=False,
                    info="チェックすると、Excelの先頭行もデータとして読み込みます。"
                )
            with gr.Row():
                include_first_column_chk_test = gr.Checkbox(
                    label="先頭列をデータとして扱う", value=False,
                    info="チェックすると、Excelの先頭列もデータとして読み込みます。"
                )
            with gr.Row():
                calc_direction_radio_test = gr.Radio(choices=["列方向", "行方向"],
                                                     label="計算対象方向", value="列方向")
            with gr.Row():
                preview_df_test = gr.DataFrame(label="データプレビュー (先頭5行)", interactive=False)
            with gr.Row():
                test_column_dropdown = gr.Dropdown(choices=[], label="検定対象の列または行を2つ選択してください", multiselect=True)
            with gr.Row():
                perform_f_test_radio = gr.Radio(
                    choices=["F検定を実施しない", "F検定を実施する"],
                    label="F検定",
                    value="F検定を実施する"
                )
                alpha_f_input = gr.Number(label="有意水準 (F検定)", value=0.05, precision=3)
            with gr.Row():
                perform_t_test_radio = gr.Radio(
                    choices=["T検定を実施しない", "対応ありt検定", "独立t検定（分散が等しい）", "独立t検定（分散が異なる）"],
                    label="t検定",
                    value="対応ありt検定"
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
                run_test_button = gr.Button("解析実行")
            with gr.Row():
                test_result_box = gr.Textbox(label="検定結果・ログ", lines=10, interactive=False)
            with gr.Row():
                density_overlay_gallery = gr.Gallery(label="理論分布プロット", show_label=True, type="file")
            with gr.Row():
                excel_test_file_box = gr.File(label="出力されたExcelファイルを開く")
                excel_test_preview_box = gr.DataFrame(label="Excelファイルの内容プレビュー", interactive=False)
            with gr.Row():
                open_folder_test_button = gr.Button("Outputフォルダを開く")
            
            test_file_input.change(
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
                    perform_t_test_radio, perform_t_test_radio, alpha_t_input,
                    include_first_row_chk_test, include_first_column_chk_test,
                    plot_overlay_radio, calc_corr_radio, calc_direction_radio_test
                ],
                outputs=[test_result_box, density_overlay_gallery, excel_test_file_box, excel_test_preview_box]
            )
            open_folder_test_button.click(fn=open_output_folder, inputs=[], outputs=[])      

        with gr.Tab("📖 初学者向け解説"):
            try:
                with open("explanation.txt", "r", encoding="utf-8") as f:
                    explanation_text = f.read()
            except Exception as e:
                explanation_text = f"解説ファイルの読み込みに失敗しました: {e}"
            gr.Markdown(explanation_text)

    version = get_version()
    gr.Markdown(f"©2025 @KotaOoka  |  **バージョン: {version}**")
    
demo.launch(inbrowser=True)

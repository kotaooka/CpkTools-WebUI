import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt
import datetime
import os
import numpy as np
import math
from PIL import Image
import gradio as gr

# -------------------------
# 初期設定：日本語フォント設定（Windowsの場合）
plt.rcParams['font.family'] = 'MS Gothic'

# -------------------------
# 出力先ディレクトリの設定（固定）
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------
# 補助関数: スペックテーブルを更新（インタラクティブに対応）
def update_spec_df_with_checkbox(selected_columns, same_spec, current_spec):
    if not selected_columns:
        return []
    if isinstance(current_spec, pd.DataFrame):
        current_spec_list = current_spec.values.tolist()
    elif current_spec is None:
        current_spec_list = []
    else:
        current_spec_list = current_spec

    new_spec = []
    for idx, col in enumerate(selected_columns):
        if idx < len(current_spec_list) and current_spec_list[idx] and current_spec_list[idx][0] == col:
            new_spec.append(current_spec_list[idx])
        else:
            new_spec.append([col, "", ""])
    if same_spec and len(new_spec) > 0:
        first_usl = new_spec[0][1]
        first_lsl = new_spec[0][2]
        new_spec = [[row[0], first_usl, first_lsl] for row in new_spec]
    return new_spec

# -------------------------
# ファイルアップロード時のプレビュー更新（先頭行の扱いを選択可能）
def update_preview(uploaded_file, include_first_row):
    if uploaded_file is None:
        return None, gr.update(choices=[])
    try:
        # 「先頭行をデータとして扱う」なら header=None として全行読み込み
        if include_first_row:
            df = pd.read_excel(uploaded_file.name, header=None)
            # 自動的に列名("Column A", "Column B", …)を付与
            df.columns = [f"Column {chr(65+i)}" for i in range(len(df.columns))]
        else:
            df = pd.read_excel(uploaded_file.name, header=0)
    except Exception as e:
        return f"ファイル読み込みエラー: {e}", gr.update(choices=[])
    if df.empty:
        return "ファイルにデータがありません。", gr.update(choices=[])
    # プレビュー＆解析対象列選択用に、各列名を「A列 (実際の列名)」の形式で表示
    column_choices = [f"{chr(65 + i)}列 ({col})" for i, col in enumerate(df.columns)]
    return df.head(5), gr.update(choices=column_choices)

# -------------------------
# 解析処理
def run_analysis(uploaded_file, selected_columns, spec_table, subgroup_size, include_first_row,
                 show_hist, show_qq, show_density, show_xbar, show_r, show_s, std_method):
    log_messages = ""
    hist_images = []      # ヒストグラム
    qq_images = []        # QQプロット
    density_images = []   # 確率密度分布
    xbar_images = []      # X-barチャート または Iチャート
    r_images = []         # Rチャート または MRチャート
    s_images = []         # s管理図（サブグループサイズ>=2の場合のみ）
    excel_file = None     # Excel出力ファイルパス
    excel_preview = None  # Excel出力結果プレビュー
    results = []          # 各列の統計解析結果リスト

    # ユーザー選択に応じた自由度の設定：
    # 「サンプル標準偏差」なら ddof=1, 「母集団標準偏差」なら ddof=0
    ddof_value = 1 if std_method == "サンプル標準偏差" else 0

    if uploaded_file is None:
        return "エラー: ファイルが選択されていません", None, None, None, None, None, None, None, None

    try:
        if include_first_row:
            df = pd.read_excel(uploaded_file.name, header=None)
            df.columns = [f"Column {chr(65+i)}" for i in range(len(df.columns))]
        else:
            df = pd.read_excel(uploaded_file.name, header=0)
        log_messages += "ファイル読み込み成功。\n"
    except Exception as e:
        return f"エラー: ファイル読み込み中に問題が発生しました: {e}", None, None, None, None, None, None, None, None

    if df.empty:
        return "エラー: ファイルにデータがありません", None, None, None, None, None, None, None, None
    if not selected_columns:
        return "エラー: 解析対象の列が選択されていません", None, None, None, None, None, None, None, None

    try:
        spec_df = pd.DataFrame(spec_table, columns=["解析対象", "規格上限値", "規格下限値"])
    except Exception as e:
        return f"エラー: 規格値テーブルの読み込みに失敗しました: {e}", None, None, None, None, None, None, None, None

    if len(spec_df) != len(selected_columns):
        return "エラー: 選択された列数と規格値入力の行数が一致しません。", None, None, None, None, None, None, None, None

    for i, col_label in enumerate(selected_columns):
        try:
            column_index = ord(col_label[0]) - 65
        except Exception as e:
            log_messages += f"エラー: 選択された列の形式が正しくありません ({col_label})\n"
            continue
        if column_index < 0 or column_index >= len(df.columns):
            log_messages += f"エラー: 正しい列を選択してください ({col_label})\n"
            continue

        actual_column = df.columns[column_index]
        data = df[actual_column]

        # 欠損値があれば除外し、注意ログを出力
        if data.isnull().any():
            log_messages += f"注意: {col_label} に欠損値が存在します。欠損値を除外して解析します。（元のデータ数: {len(data)}）\n"
            data = data.dropna()

        sample_n = len(data)
        if sample_n == 0:
            log_messages += f"エラー: {col_label} のデータがすべて欠損しています。\n"
            continue

        try:
            spec_entry_usl = spec_df.iloc[i]["規格上限値"]
            spec_entry_lsl = spec_df.iloc[i]["規格下限値"]
            current_usl = float(spec_entry_usl) if str(spec_entry_usl).strip() != "" else None
            current_lsl = float(spec_entry_lsl) if str(spec_entry_lsl).strip() != "" else None
        except Exception as e:
            log_messages += f"エラー: {col_label} の規格値が正しく入力されていません: {e}\n"
            continue

        try:
            max_val = float(data.max())
            min_val = float(data.min())
            # ユーザー選択に応じた標準偏差計算
            std_val = float(data.std(ddof=ddof_value))
            mean_val = float(data.mean())
            kurtosis_val = float(data.kurtosis())
            skewness_val = float(data.skew())
            if std_val == 0:
                log_messages += f"エラー: {col_label} の標準偏差が0のため、Cp/Cpk計算をスキップしました。（サンプル数: {sample_n}）\n"
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
                log_messages += f"エラー: {col_label} の規格値が入力されていません。\n"
                continue

            results.append({
                "解析対象": col_label,
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
            log_messages += f"解析対象: {col_label} ({actual_column}) の統計計算完了。（サンプル数: {sample_n}）\n"
        except Exception as e:
            log_messages += f"エラー: {col_label} の統計計算中に問題が発生しました: {e}\n"
            continue

        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')

        # ヒストグラムの生成
        if show_hist:
            try:
                plt.figure()
                plt.hist(data, color="skyblue", edgecolor="black")
                plt.xlabel("値")
                plt.ylabel("度数")
                plt.title(f"ヒストグラム ({col_label})")
                hist_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_hist_{col_label}.jpg")
                plt.savefig(hist_filename, format="jpg")
                plt.close()
                hist_images.append(hist_filename)
                log_messages += f"{col_label} のヒストグラム生成完了。\n"
            except Exception as e:
                log_messages += f"エラー: {col_label} のヒストグラム生成中に問題が発生しました: {e}\n"

        # QQプロットの生成
        if show_qq:
            try:
                plt.figure()
                stats.probplot(data, dist="norm", plot=plt)
                plt.title(f"QQプロット ({col_label})")
                qq_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_qq_{col_label}.jpg")
                plt.savefig(qq_filename, format="jpg")
                plt.close()
                qq_images.append(qq_filename)
                log_messages += f"{col_label} のQQプロット生成完了。\n"
            except Exception as e:
                log_messages += f"エラー: {col_label} のQQプロット生成中に問題が発生しました: {e}\n"

        # 確率密度分布の生成
        if show_density:
            try:
                plt.figure()
                x = np.linspace(mean_val - 4 * std_val, mean_val + 4 * std_val, 100)
                y = stats.norm.pdf(x, loc=mean_val, scale=std_val)
                plt.plot(x, y, label="正規分布", color="blue")
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
                plt.title(f"確率密度分布 ({col_label})")
                plt.legend()
                density_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_density_{col_label}.jpg")
                plt.savefig(density_filename, format="jpg")
                plt.close()
                density_images.append(density_filename)
                log_messages += f"{col_label} の確率密度分布描画完了。\n"
            except Exception as e:
                log_messages += f"エラー: {col_label} の確率密度分布描画中に問題が発生しました: {e}\n"

        # --- サブグループチャートの生成 ---
        if subgroup_size == 1:
            individuals = data.values
            n_individuals = len(individuals)
            if n_individuals < 1:
                log_messages += f"警告: {col_label} のデータ点数が不足しているため、チャートを生成できませんでした。\n"
            else:
                # Iチャートの計算
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
                        plt.title(f"Iチャート ({col_label})")
                        plt.legend()
                        i_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_i_{col_label}.jpg")
                        plt.savefig(i_filename, format="jpg")
                        plt.close()
                        xbar_images.append(i_filename)
                        log_messages += f"{col_label} のIチャート生成完了。\n"
                    except Exception as e:
                        log_messages += f"エラー: {col_label} のIチャート生成中に問題が発生しました: {e}\n"
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
                            plt.title(f"MRチャート ({col_label})")
                            plt.legend()
                            mr_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_mr_{col_label}.jpg")
                            plt.savefig(mr_filename, format="jpg")
                            plt.close()
                            r_images.append(mr_filename)
                            log_messages += f"{col_label} のMRチャート生成完了。\n"
                        except Exception as e:
                            log_messages += f"エラー: {col_label} のMRチャート生成中に問題が発生しました: {e}\n"
                    else:
                        log_messages += f"警告: {col_label} のデータ点数が不十分なため、MRチャートを生成できませんでした。\n"
                if show_s:
                    log_messages += f"警告: サブグループサイズが1のため、s管理図は生成できません。\n"
        elif subgroup_size >= 2 and (show_xbar or show_r or show_s):
            try:
                if len(data) >= subgroup_size:
                    n_groups = int(np.ceil(len(data) / subgroup_size))
                    subgroup_means = []
                    subgroup_ranges = []
                    subgroup_stds = []
                    for j in range(n_groups):
                        subgroup = data.iloc[j * subgroup_size : min((j + 1) * subgroup_size, len(data))]
                        subgroup_means.append(np.mean(subgroup))
                        subgroup_ranges.append(np.max(subgroup) - np.min(subgroup))
                        if len(subgroup) >= 2:
                            # ddof_value を用いてサブグループ標準偏差を計算
                            subgroup_stds.append(np.std(subgroup, ddof=ddof_value))
                    xbar_bar = np.mean(subgroup_means)
                    R_bar = np.mean(subgroup_ranges)
                    chart_factors = {
                        2: {"A2": 1.88, "D3": 0.0,   "D4": 3.267},
                        3: {"A2": 1.023, "D3": 0.0,  "D4": 2.574},
                        4: {"A2": 0.729, "D3": 0.0,  "D4": 2.282},
                        5: {"A2": 0.577, "D3": 0.0,  "D4": 2.114},
                        6: {"A2": 0.483, "D3": 0.0,  "D4": 2.004},
                        7: {"A2": 0.419, "D3": 0.076,"D4": 1.924},
                        8: {"A2": 0.373, "D3": 0.136,"D4": 1.864},
                        9: {"A2": 0.337, "D3": 0.184,"D4": 1.816},
                        10: {"A2": 0.308, "D3": 0.223,"D4": 1.777},
                    }
                    if subgroup_size in chart_factors:
                        A2 = chart_factors[subgroup_size]["A2"]
                        D3 = chart_factors[subgroup_size]["D3"]
                        D4 = chart_factors[subgroup_size]["D4"]

                        if show_xbar:
                            plt.figure()
                            plt.plot(range(1, n_groups + 1), subgroup_means, marker='o', linestyle='-', color='blue', label='サブグループ平均')
                            plt.axhline(xbar_bar, color='green', linestyle='--', label='全体平均')
                            plt.axhline(xbar_bar + A2 * R_bar, color='red', linestyle='--', label='上限管理限界')
                            plt.axhline(xbar_bar - A2 * R_bar, color='red', linestyle='--', label='下限管理限界')
                            if current_usl is not None:
                                plt.axhline(current_usl, color='magenta', linestyle='-.', label='規格上限値')
                            if current_lsl is not None:
                                plt.axhline(current_lsl, color='cyan', linestyle='-.', label='規格下限値')
                            plt.xlabel('サブグループ')
                            plt.ylabel('平均値')
                            plt.title(f"X-barチャート ({col_label})")
                            plt.legend()
                            xbar_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_xbar_{col_label}.jpg")
                            plt.savefig(xbar_filename, format="jpg")
                            plt.close()
                            xbar_images.append(xbar_filename)
                            log_messages += f"{col_label} のX-barチャート生成完了。\n"

                        if show_r:
                            plt.figure()
                            plt.plot(range(1, n_groups + 1), subgroup_ranges, marker='o', linestyle='-', color='blue', label='サブグループレンジ')
                            plt.axhline(R_bar, color='green', linestyle='--', label='平均レンジ')
                            plt.axhline(D4 * R_bar, color='red', linestyle='--', label='UCL')
                            plt.axhline(D3 * R_bar, color='red', linestyle='--', label='LCL')
                            plt.xlabel('サブグループ')
                            plt.ylabel('レンジ')
                            plt.title(f"Rチャート ({col_label})")
                            plt.legend()
                            r_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_r_{col_label}.jpg")
                            plt.savefig(r_filename, format="jpg")
                            plt.close()
                            r_images.append(r_filename)
                            log_messages += f"{col_label} のRチャート生成完了。\n"

                        if show_s and subgroup_stds:
                            s_bar = np.mean(subgroup_stds)
                            c4 = math.sqrt(2/(subgroup_size-1)) * math.exp(math.lgamma(subgroup_size/2) - math.lgamma((subgroup_size-1)/2))
                            sigma_s = s_bar * math.sqrt(1 - c4**2) / c4
                            UCL_s = s_bar + 3 * sigma_s
                            LCL_s = s_bar - 3 * sigma_s
                            if LCL_s < 0:
                                LCL_s = 0

                            plt.figure()
                            plt.plot(range(1, n_groups+1), subgroup_stds, marker='o', linestyle='-', color='blue', label='サブグループ標準偏差')
                            plt.axhline(s_bar, color='green', linestyle='--', label='全体平均標準偏差')
                            plt.axhline(UCL_s, color='red', linestyle='--', label='UCL')
                            plt.axhline(LCL_s, color='red', linestyle='--', label='LCL')
                            plt.xlabel('サブグループ')
                            plt.ylabel('標準偏差')
                            plt.title(f"s管理図 ({col_label})")
                            plt.legend()
                            s_filename = os.path.join(OUTPUT_DIR, f"{timestamp}_s_{col_label}.jpg")
                            plt.savefig(s_filename, format="jpg")
                            plt.close()
                            s_images.append(s_filename)
                            log_messages += f"{col_label} のs管理図生成完了。\n"
                        else:
                            if show_s:
                                log_messages += f"警告: {col_label} のサブグループ標準偏差の計算に十分なデータがないため、s管理図を生成できませんでした。\n"
                    else:
                        log_messages += f"警告: サブグループサイズ {subgroup_size} に対するチャートファクターが見つからなかったため、X-barチャートとRチャートをスキップします。\n"
                else:
                    log_messages += f"警告: {col_label} のデータ点数がサブグループサイズより少ないため、X-barチャート、Rチャートおよびs管理図を生成できませんでした。\n"
            except Exception as e:
                log_messages += f"エラー: {col_label} のX-bar/R/sチャート生成中に問題が発生しました: {e}\n"

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
        log_messages += "エラー: 解析対象の列から有効なデータが得られませんでした。\n"

    return log_messages, hist_images, qq_images, density_images, xbar_images, r_images, s_images, excel_file, excel_preview

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
        with gr.Tab("📊 解析ツール"):
            with gr.Row():
                file_input = gr.File(label="Excelファイル (xlsx, xls)", file_count="single")
            with gr.Row():
                include_first_row_chk = gr.Checkbox(
                    label="先頭行をデータとして扱う", value=False,
                    info="チェックすると、Excelの先頭行もデータとして読み込みます。"
                )
            with gr.Row():
                preview_df = gr.DataFrame(label="データプレビュー (先頭5行)", interactive=False)
            with gr.Row():
                column_dropdown = gr.Dropdown(choices=[], label="解析対象の列 (A列, B列, ...)", multiselect=True)
            with gr.Row():
                spec_df = gr.Dataframe(
                    headers=["解析対象", "規格上限値", "規格下限値"],
                    label="各列の規格値入力（空欄は片側規格として扱います）", interactive=True
                )
            with gr.Row():
                same_spec_chk = gr.Checkbox(label="すべての列の規格値を同じにする", value=False)
            with gr.Row():
                subgroup_size_slider = gr.Slider(
                    minimum=1, maximum=10, step=1, value=5,
                    label="サブグループサイズ",
                    info="X-barチャート、Rチャート、s管理図作成時に用いるサブグループのサイズ。1の場合はIチャート/MRチャートを生成します。"
                )
            with gr.Row():
                show_hist_checkbox = gr.Checkbox(label="ヒストグラムを出力", value=True)
                show_qq_checkbox = gr.Checkbox(label="QQプロットを出力", value=True)
                show_density_checkbox = gr.Checkbox(label="確率密度分布を出力", value=True)
            with gr.Row():
                show_xbar_checkbox = gr.Checkbox(label="X-barチャート／Iチャートを出力", value=True)
                show_r_checkbox = gr.Checkbox(label="Rチャート／MRチャートを出力", value=True)
                show_s_checkbox = gr.Checkbox(label="s管理図を出力", value=True)
            with gr.Row():
                # 新たに標準偏差の計算方法を選択するラジオボタンを追加
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
                xbar_gallery = gr.Gallery(label="X-barチャート／Iチャート", show_label=True, type="file")
                r_gallery = gr.Gallery(label="Rチャート／MRチャート", show_label=True, type="file")
            with gr.Row():
                s_gallery = gr.Gallery(label="s管理図", show_label=True, type="file")
            with gr.Row():
                excel_file_box = gr.File(label="出力されたExcelファイルを開く")
                excel_preview_box = gr.DataFrame(label="Excelファイルの内容プレビュー", interactive=False)
            with gr.Row():
                open_folder_button = gr.Button("Outputフォルダを開く")
            
            file_input.change(
                fn=update_preview,
                inputs=[file_input, include_first_row_chk],
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
                    file_input, column_dropdown, spec_df, subgroup_size_slider, include_first_row_chk,
                    show_hist_checkbox, show_qq_checkbox, show_density_checkbox,
                    show_xbar_checkbox, show_r_checkbox, show_s_checkbox, std_method_radio
                ],
                outputs=[
                    result_box, hist_gallery, qq_gallery, density_gallery,
                    xbar_gallery, r_gallery, s_gallery, excel_file_box, excel_preview_box
                ]
            )
            open_folder_button.click(fn=open_output_folder, inputs=[], outputs=[])
        with gr.Tab("📖 初学者向け解説"):
            gr.Markdown(
                """
        # 品質管理初学者向け解説ページ

        本解説ページでは、工程能力解析ツールで算出される各統計量やグラフの意味、計算方法、各グラフの見方、そして実務上の活用ポイントについて体系的に解説します。  
        以下の各項目を順に学ぶことで、工程の状態の把握と改善に役立つ知識を得ることができます。

        ---

        ## 1. 基本統計量と分布の特徴

        ### 1.1 平均値と標準偏差
        - **平均値**:  
        工程全体の中心的な値を示し、代表値として利用されます。
        - **標準偏差**:  
        各データ点が平均値からどれだけ離れているかを数値化したもので、工程のばらつきを示します。  
        標準偏差が小さいほど、データは平均値周辺に集中しています。

        ### 1.2 尖度と歪度
        - **歪度 (Skewness)**:  
        分布の左右対称性を評価する指標です。  
        - |歪度| < 0.5：ほぼ対称  
        - 0.5 ≤ |歪度| < 1.0：中程度の偏り  
        - |歪度| ≥ 1.0：顕著な偏りがある
        - **尖度 (Kurtosis)**:  
        分布のピークの鋭さや尾部の重さ（外れ値の出やすさ）を示します（余剰尖度の場合、正規分布は0が基準）。  
        - -0.5 ～ 0.5：正規分布に近い  
        - 0.5 ～ 1 または -0.5 ～ -1：やや尖った平らな分布  
        - |尖度| > 1：非常に重い尾や平らなピークで、外れ値の影響が大きい

        これらの基本統計量は、まずヒストグラムやQQプロットを用いてデータの正規性を確認する際の基礎となります。

        ---

        ## 2. 各グラフの見方

        ### 2.1 ヒストグラム
        - **目的**:  
        データの度数分布や集中の傾向、外れ値の有無を直感的に把握するためのグラフです。
        - **見方**:  
        - 横軸: 測定値の範囲  
        - 縦軸: 各範囲に該当するデータ数（度数）
        - **ポイント**:  
        - どの範囲にデータが密集しているか  
        - 双峰性（2つ以上のピーク）があるか  
        - 右裾または左裾の伸び具合で分布の非対称性を示すか

        ### 2.2 QQプロット
        - **目的**:  
        データが正規分布に従っているかを視覚的に評価するためのプロットです。
        - **見方**:  
        予測される正規分布の理論分位点と実測分位点がプロットされ、理想的には直線上に並びます。
        - **ポイント**:  
        点が直線から大きく外れる場合、正規性に逸脱が見られる（外れ値の影響も示唆）。

        ### 2.3 確率密度分布（Density Plot）
        - **目的**:  
        正規分布カーブと実際のデータ分布を比較し、平均値や±3σの位置に対するデータの分布状況を確認します。
        - **見方**:  
        - カーブの形状やピークの位置、裾の伸び具合を確認  
        - 規格値（USL, LSL）の位置と重なり具合をチェック
        - **ポイント**:  
        平均値、±3σの位置がどのように評価されるかで、工程のリスク評価に役立ちます。

        ### 2.4 X-barチャートとIチャート

        #### X-barチャート
        - **目的**:  
        複数の測定値から算出されたサブグループの平均値を時系列でプロットし、工程の中心位置とばらつきを監視します。
        - **見方**:  
        - 横軸: サブグループ番号  
        - 縦軸: 各サブグループの平均値  
        - 管理限界（上限・下限）は、全体平均とサブグループ内のばらつき（A2係数を用いて算出）で決定される。
        - **ポイント**:  
        - 大部分の点が管理限界内にあるか  
        - 連続した偏りや急激な変動の兆候がないか

        #### Iチャート (Individuals Chart)
        - **目的**:  
        個々の測定値をそのままプロットし、リアルタイムの工程変動を把握します。
        - **見方**:  
        - 横軸: 各データ取得の順序  
        - 縦軸: 各測定値  
        - 隣接するデータ点の変動（移動範囲：MRチャート）と連動して評価される。
        - **ポイント**:  
        - 外れ値や急激な変動が確認できるか  
        - リアルタイム性の高いデータ監視に適しているが、外れ値の影響を受けやすい

        ### 2.5 Rチャート / MRチャート
        - **目的**:  
        サブグループや隣接するデータの範囲（または移動範囲）から、工程内のばらつきを定量的に評価します。
        - **見方**:  
        - 横軸: サブグループ番号またはデータ点の序列  
        - 縦軸: 各サブグループのレンジまたは移動範囲
        - **ポイント**:  
        レンジの変動が小さいか、急激な変動点がある場合は工程の不安定が疑われます。

        ### 2.6 s管理図
        - **目的**:  
        各サブグループの標準偏差をプロットし、工程のばらつきの変化をモニタリングします。
        - **見方**:  
        - 横軸: サブグループ番号  
        - 縦軸: 各サブグループの標準偏差
        - **ポイント**:  
        管理限界との比較により、ばらつきが一定かどうかを評価します。  
        大きな変動があれば工程の改善点となります。

        ---

        ## 3. 工程能力指数（Cp, Cpk）と不良率の関係

        ### 3.1 Cp と Cpk の概要
        - **Cp**:  
        **定義**: 規格幅に対する工程変動の小ささを評価します。  
        ```
        Cp = (規格上限値 - 規格下限値) / (6 * 標準偏差)
        ```
        - **Cpk**:  
        **定義**: 工程のばらつきに加えて、工程の平均（中心）が仕様範囲内のどちらかの限界（USLまたはLSL）からどれだけずれているかも考慮し、実際の工程能力を評価する指標です。
        ```
        Cpk = min((規格上限値 - 平均値) / (3 * 標準偏差), (平均値 - 規格下限値) / (3 * 標準偏差))
        ```
        - **解釈**:  
        一般に Cpk が 1.33 以上であれば、工程は十分な能力を持つとされます。

        ### 3.2 不良率の計算式
        - **両側規格の場合 (Cp1)**  
        正規分布を前提とすると、平均値から±3σ内に約99.73%のデータが含まれるため、不良率は 
        ```
        不良率 = 2 × (1 - Φ(3)) ≈ 0.27%
        ```
        ここで、Φ(3)は標準正規分布における平均から3σまでの累積確率を示し、その値はおよそ0.99865です。すなわち、
        ```
        2×(1−0.99865)≈0.27%
        ```
        この計算式により、両側規格の場合の工程内での不良品の発生率がおおよそ0.27%であると導かれます。

        - **片側規格の場合 (Cp1)**
        - **Cp1の計算式 (上側規格のみの例)**  
        ```
        Cp1 = (規格上限値 - 平均値) / (3 × σ)
        ```
        - この場合、不良率は  
        ```
        不良率 = 1 - Φ(3) ≈ 0.135%
        ```
        ※ これらは理論上の値であり、実際の工程ではプロセスの偏りや非正規性により変動します。

        ---

        ## 4. サブグループサイズとA2係数の関係

        ### 4.1 A2係数の役割と算出方法
        - **A2係数**は、X-barチャートでサブグループごとの平均レンジから管理限界を設定するための係数です。  
        管理限界は次の式で計算されます。
        ```
        管理限界 = 全体平均 ± (A2 × サブグループ平均レンジ)
        ```
        ### 4.2 サブグループサイズとA2係数の関係

        | サブグループサイズ | A2係数  |
        | ------------------ | ------- |
        | 2                  | 1.88    |
        | 3                  | 1.023   |
        | 4                  | 0.729   |
        | 5                  | 0.577   |
        | 6                  | 0.483   |
        | 7                  | 0.419   |
        | 8                  | 0.373   |
        | 9                  | 0.337   |
        | 10                 | 0.308   |

        ### 4.3 解説
        - **小さいサブグループ（例：サイズ2～3）**:  
        各グループ内のばらつきが大きく反映されるため、A2係数が高くなり、管理限界は広く設定されます。
        - **大きいサブグループ（例：サイズ5以上）**:  
        サブグループの平均が安定するため、A2係数は低下し、管理限界が狭くなります。
        - **実務上の注意**:  
        適切なサブグループサイズの選択は、偽陽性の警報を防ぎ、実際の工程異常を正確に捉えるうえで重要です。

        ---

        ## 5. 実務上の注意点とまとめ

        - **正規性の確認**:  
        Cp/Cpk/Cp1 の計算は正規分布の前提に依存するため、まずヒストグラム、QQプロット、確率密度分布でデータの正規性を確認してください。

        - **サンプルサイズの確保**:  
        少ないデータの場合、統計量が不安定になるため、十分なサンプル数を確保する必要があります。

        - **工程の中心性への注意**:  
        Cpk は工程の平均の位置も評価するため、偏りがある場合はその要因の是正を検討してください。

        - **管理図の総合活用**:  
        X-barチャート、Iチャート、R/MRチャート、s管理図を併用することで、工程の状態を多角的に評価し、早期に異常を発見することが可能です。

        ### まとめ
        - 基本統計量（平均、標準偏差、尖度、歪度）と工程能力指数（Cp、Cpk、Cp1）を理解し、それぞれの評価指標と理論値を把握することが重要です。  
        - 各グラフの見方（ヒストグラム、QQプロット、確率密度分布、各種管理図）を理解した上で、実測データと理論値を比較し、工程の安定性や改善すべき箇所を判断します。  
        - 特に片側規格の場合のCp1は、理論上約1350 ppmの不良率が期待されるものの、実際には工程の偏りや分布の非正規性を考慮して運用する必要があります。  
        - サブグループサイズとA2係数の関係を正しく把握することで、X-barチャートによる管理限界の設定と異常検出の精度が向上します.
            """
        )



    gr.Markdown("©2025 @KotaOoka")
    
demo.launch(inbrowser=True)

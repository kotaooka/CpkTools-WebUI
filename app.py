import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt
import datetime
import os
import io
import json
import numpy as np  # numpy のインポート追加

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
# 補助関数: 
# 「すべての列の規格値を同じにする」チェックボックスに基づいて、規格値入力テーブルを更新する
def update_spec_df_with_checkbox(selected_columns, same_spec, current_spec):
    if not selected_columns:
        return []
    # current_spec が DataFrame ならリストに変換、それ以外はリストとして扱う
    if isinstance(current_spec, pd.DataFrame):
        current_spec_list = current_spec.values.tolist()
    elif current_spec is None:
        current_spec_list = []
    else:
        current_spec_list = current_spec

    # 既存の入力内容を出来るだけ保持し、選択された列に対応する行を作成
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
# ファイルアップロード時のプレビュー更新
def update_preview(uploaded_file):
    if uploaded_file is None:
        return None, gr.update(choices=[])
    try:
        df = pd.read_excel(uploaded_file.name)
    except Exception as e:
        return f"ファイル読み込みエラー: {e}", gr.update(choices=[])
    if df.empty:
        return "ファイルにデータがありません。", gr.update(choices=[])
    column_choices = [f"{chr(65 + i)}列 ({col})" for i, col in enumerate(df.columns)]
    return df.head(5), gr.update(choices=column_choices)

# -------------------------
# 解析処理：各列の統計計算、グラフ生成、Excel出力およびプレビュー表示
def run_analysis(uploaded_file, selected_columns, spec_table):
    log_messages = ""
    hist_images = []      # ヒストグラム画像のリスト
    qq_images = []        # QQプロット画像のリスト
    density_images = []   # 確率密度分布画像のリスト
    excel_file = None     # 出力した Excel ファイルのパス
    excel_preview = None  # 統計結果の DataFrame

    if uploaded_file is None:
        return "エラー: ファイルが選択されていません", None, None, None, None, None
    try:
        df = pd.read_excel(uploaded_file.name)
        log_messages += "ファイル読み込み成功。\n"
    except Exception as e:
        return f"エラー: ファイル読み込み中に問題が発生しました: {e}", None, None, None, None, None
    if df.empty:
        return "エラー: ファイルにデータがありません", None, None, None, None, None
    if not selected_columns:
        return "エラー: 解析対象の列が選択されていません", None, None, None, None, None

    try:
        spec_df = pd.DataFrame(spec_table, columns=["解析対象", "規格上限値", "規格下限値"])
    except Exception as e:
        return f"エラー: 規格値テーブルの読み込みに失敗しました: {e}", None, None, None, None, None
    if len(spec_df) != len(selected_columns):
        return "エラー: 選択された列数と規格値入力の行数が一致しません。", None, None, None, None, None

    results = []
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
        if data.isnull().any():
            log_messages += f"エラー: {col_label} に欠損値があります\n"
            continue
        try:
            current_usl = float(spec_df.iloc[i]["規格上限値"])
            current_lsl = float(spec_df.iloc[i]["規格下限値"])
        except Exception as e:
            log_messages += f"エラー: {col_label} の規格値が正しく入力されていません: {e}\n"
            continue

        try:
            max_val = float(data.max())
            min_val = float(data.min())
            std_val = float(data.std())
            mean_val = float(data.mean())
            kurtosis_val = float(data.kurtosis())
            skewness_val = float(data.skew())
            if std_val == 0:
                log_messages += f"エラー: {col_label} の標準偏差が0のため、CpおよびCpkの計算をスキップしました。\n"
                continue
            Cp = (current_usl - current_lsl) / (6 * std_val)
            Cpk = min((current_usl - mean_val), (mean_val - current_lsl)) / (3 * std_val)
            results.append({
                "解析対象": col_label,
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
            log_messages += f"解析対象: {col_label} ({actual_column}) の統計計算完了。\n"
        except Exception as e:
            log_messages += f"エラー: {col_label} の統計計算中に問題が発生しました: {e}\n"
            continue

        # ヒストグラム図の生成
        try:
            plt.figure()
            plt.hist(data, color="skyblue", edgecolor="black")
            plt.xlabel("値")
            plt.ylabel("度数")
            plt.title(f"ヒストグラム ({col_label})")
            buf_hist = io.BytesIO()
            plt.savefig(buf_hist, format="png")
            plt.close()
            buf_hist.seek(0)
            image_hist = Image.open(buf_hist)
            hist_images.append(image_hist)
            log_messages += f"{col_label} のヒストグラム生成完了。\n"
        except Exception as e:
            log_messages += f"エラー: {col_label} のヒストグラム生成中に問題が発生しました: {e}\n"

        # QQプロット図の生成
        try:
            plt.figure()
            stats.probplot(data, dist="norm", plot=plt)
            plt.title(f"QQプロット ({col_label})")
            buf_qq = io.BytesIO()
            plt.savefig(buf_qq, format="png")
            plt.close()
            buf_qq.seek(0)
            image_qq = Image.open(buf_qq)
            qq_images.append(image_qq)
            log_messages += f"{col_label} のQQプロット生成完了。\n"
        except Exception as e:
            log_messages += f"エラー: {col_label} のQQプロット生成中に問題が発生しました: {e}\n"

        # 確率密度分布図の生成（±3σ、平均値、規格上限値・下限値のラインとラベル）
        try:
            plt.figure()
            # ±4σの範囲でx軸を作成（±3σが含まれるように）
            x = np.linspace(mean_val - 4 * std_val, mean_val + 4 * std_val, 100)
            y = stats.norm.pdf(x, loc=mean_val, scale=std_val)
            plt.plot(x, y, label="正規分布", color="blue")
            
            # ±3σ のラインの追加
            plt.axvline(mean_val - 3 * std_val, color="red", linestyle="--", label="-3σ")
            plt.axvline(mean_val + 3 * std_val, color="red", linestyle="--", label="+3σ")
            # 平均値のライン（橙色の実線）
            plt.axvline(mean_val, color="orange", linestyle="-", label="平均値")
            # 規格上限値（緑色の点線）と規格下限値（紫色の点線）のラインを追加
            plt.axvline(current_usl, color="green", linestyle="-.", label="規格上限値")
            plt.axvline(current_lsl, color="purple", linestyle="-.", label="規格下限値")
            
            # ラベルを追加するために現在のAxesを取得
            ax = plt.gca()
            # グラフ上部のy座標(現在のy軸上限値)を取得し、少し下げた位置にラベルを配置する
            y_top = ax.get_ylim()[1]
            label_y = y_top * 0.20
            
            # 各ライン上に値のラベル（数値を整形して表示）
            ax.text(mean_val - 3 * std_val, label_y, f"-3σ: {mean_val - 3 * std_val:.2f}", 
                    rotation=90, color="black", ha="center", va="bottom", fontsize=8)
            ax.text(mean_val + 3 * std_val, label_y, f"+3σ: {mean_val + 3 * std_val:.2f}", 
                    rotation=90, color="black", ha="center", va="bottom", fontsize=8)
            ax.text(mean_val, label_y, f"平均値: {mean_val:.2f}", 
                    rotation=90, color="black", ha="center", va="bottom", fontsize=8)
            ax.text(current_usl, label_y, f"規格上限値: {current_usl:.2f}", 
                    rotation=90, color="black", ha="center", va="bottom", fontsize=8)
            ax.text(current_lsl, label_y, f"規格下限値: {current_lsl:.2f}", 
                    rotation=90, color="black", ha="center", va="bottom", fontsize=8)
            
            plt.xlabel("値")
            plt.ylabel("確率密度")
            plt.title(f"確率密度分布 ({col_label})")
            plt.legend()
            buf_pdf = io.BytesIO()
            plt.savefig(buf_pdf, format="png")
            plt.close()
            buf_pdf.seek(0)
            image_pdf = Image.open(buf_pdf)
            density_images.append(image_pdf)
            log_messages += f"{col_label} の確率密度分布描画完了。\n"
        except Exception as e:
            log_messages += f"エラー: {col_label} の確率密度分布描画中に問題が発生しました: {e}\n"

    if results:
        dt_now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = os.path.join(OUTPUT_DIR, f"{dt_now}_results.xlsx")
        try:
            df_result = pd.DataFrame(results)
            df_result.to_excel(output_filename, index=False)
            excel_file = output_filename
            excel_preview = df_result
            log_messages += f"Excelファイル出力完了: {output_filename}\n"
        except Exception as e:
            log_messages += f"エラー: Excelファイル書き出し中に問題が発生しました: {e}\n"
    else:
        log_messages += "エラー: 解析対象の列から有効なデータが得られませんでした。\n"

    return log_messages, hist_images, qq_images, density_images, excel_file, excel_preview

# -------------------------
# Gradio UI の構築
with gr.Blocks() as demo:
    gr.Markdown("# 🏭 CpkTools-WebUI 工程能力解析ツール")
    
    with gr.Tab("📊 解析ツール"):
        with gr.Row():
            file_input = gr.File(label="Excelファイル (xlsx, xls)", file_count="single")
        with gr.Row():
            preview_df = gr.DataFrame(label="データプレビュー (先頭5行)", interactive=False)
        with gr.Row():
            column_dropdown = gr.Dropdown(choices=[], label="解析対象の列 (A列, B列, ...)", multiselect=True)
        with gr.Row():
            spec_df = gr.Dataframe(headers=["解析対象", "規格上限値", "規格下限値"],
                                   label="各列の規格値入力", interactive=True)
        with gr.Row():
            same_spec_chk = gr.Checkbox(label="すべての列の規格値を同じにする", value=False)
        run_button = gr.Button("解析開始")
        result_box = gr.Textbox(label="計算結果・ログ", lines=10, interactive=False)
        with gr.Row():
            hist_gallery = gr.Gallery(label="ヒストグラム", show_label=True)
            qq_gallery = gr.Gallery(label="QQプロット", show_label=True)
        with gr.Row():
            density_gallery = gr.Gallery(label="確率密度分布", show_label=True)
        with gr.Row():
            excel_file_box = gr.File(label="出力されたExcelファイルを開く")
            excel_preview_box = gr.DataFrame(label="Excelファイルの内容プレビュー", interactive=False)
        
        file_input.change(fn=update_preview, inputs=file_input, outputs=[preview_df, column_dropdown])
        column_dropdown.change(fn=update_spec_df_with_checkbox, 
                               inputs=[column_dropdown, same_spec_chk, spec_df],
                               outputs=spec_df)
        same_spec_chk.change(fn=update_spec_df_with_checkbox, 
                             inputs=[column_dropdown, same_spec_chk, spec_df],
                             outputs=spec_df)
        run_button.click(
            fn=run_analysis, 
            inputs=[file_input, column_dropdown, spec_df],
            outputs=[result_box, hist_gallery, qq_gallery, density_gallery, excel_file_box, excel_preview_box]
        )
    
    gr.Markdown("©2025 @KotaOoka")
    
demo.launch(inbrowser=True)

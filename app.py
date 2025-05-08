import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt
import datetime
import os
import io
import json

from PIL import Image
import gradio as gr

# -------------------------
# 初期設定：日本語フォント設定（Windowsの場合）
plt.rcParams['font.family'] = 'MS Gothic'

# -------------------------
# 設定ファイルと出力先の設定
SETTINGS_FILE = "settings.json"
DEFAULT_OUTPUT_DIR = "output"

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"output_dir": DEFAULT_OUTPUT_DIR}

def save_settings(output_dir):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump({"output_dir": output_dir}, f, indent=4, ensure_ascii=False)

def update_output_dir(new_dir):
    # 設定の保存と、出力先フォルダの存在確認・作成
    save_settings(new_dir)
    os.makedirs(new_dir, exist_ok=True)
    return f"出力先ディレクトリが変更されました: {new_dir}"

settings = load_settings()
OUTPUT_DIR = settings.get("output_dir", DEFAULT_OUTPUT_DIR)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------
# 補助関数: 
# 「すべての列の規格値を同じにする」チェックボックスに基づいて、規格値入力テーブルを更新する
def update_spec_df_with_checkbox(selected_columns, same_spec, current_spec):
    """
    selected_columns: 選択された列のリスト（例: ["A列 (価格)", "B列 (重さ)"]）
    same_spec: チェックボックスの状態（True の場合、1列目の規格値を全てにコピー）
    current_spec: 現在の規格値入力テーブル（gr.Dataframe から渡される値、DataFrame またはリスト）

    チェックボックス ON の場合、1行目の「規格上限値」「規格下限値」を全行に反映して返します。
    """
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
    """
    アップロードされた Excel ファイルの先頭5行をプレビューとして返すとともに、
    列名を「A列, B列, ...」形式に更新する。
    """
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
    hist_images = []   # ヒストグラム画像のリスト
    qq_images = []     # QQプロット画像のリスト
    excel_file = None  # 出力した Excel ファイルのパス
    excel_preview = None  # 統計結果の DataFrame

    if uploaded_file is None:
        return "エラー: ファイルが選択されていません", None, None, None, None
    try:
        df = pd.read_excel(uploaded_file.name)
        log_messages += "ファイル読み込み成功。\n"
    except Exception as e:
        return f"エラー: ファイル読み込み中に問題が発生しました: {e}", None, None, None, None
    if df.empty:
        return "エラー: ファイルにデータがありません", None, None, None, None
    if not selected_columns:
        return "エラー: 解析対象の列が選択されていません", None, None, None, None

    try:
        spec_df = pd.DataFrame(spec_table, columns=["解析対象", "規格上限値", "規格下限値"])
    except Exception as e:
        return f"エラー: 規格値テーブルの読み込みに失敗しました: {e}", None, None, None, None
    if len(spec_df) != len(selected_columns):
        return "エラー: 選択された列数と規格値入力の行数が一致しません。", None, None, None, None

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

    return log_messages, hist_images, qq_images, excel_file, excel_preview

# -------------------------
# Gradio UI の構築
with gr.Blocks() as demo:
    gr.Markdown("# 🏭 CpkTools-WebUI 工程能力解析ツール")
    
    with gr.Tabs():
        with gr.TabItem("📊 解析ツール"):
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
                excel_file_box = gr.File(label="出力されたExcelファイルを開く")
                excel_preview_box = gr.DataFrame(label="Excelファイルの内容プレビュー", interactive=False)
            
            file_input.change(fn=update_preview, inputs=file_input, outputs=[preview_df, column_dropdown])
            # 列変更時およびチェックボックス変更時に、現在のspec_dfの内容も受け取り更新
            column_dropdown.change(fn=update_spec_df_with_checkbox, 
                                   inputs=[column_dropdown, same_spec_chk, spec_df],
                                   outputs=spec_df)
            same_spec_chk.change(fn=update_spec_df_with_checkbox, 
                                 inputs=[column_dropdown, same_spec_chk, spec_df],
                                 outputs=spec_df)
            run_button.click(
                fn=run_analysis, 
                inputs=[file_input, column_dropdown, spec_df],
                outputs=[result_box, hist_gallery, qq_gallery, excel_file_box, excel_preview_box]
            )
        
        with gr.TabItem("⚙️ 設定"):
            gr.Markdown("出力先フォルダを変更できます。")
            output_dir_box = gr.Textbox(label="出力先ディレクトリ", value=OUTPUT_DIR)
            save_button = gr.Button("保存")
            setting_result = gr.Textbox(label="設定結果", lines=2)
            save_button.click(fn=update_output_dir, inputs=[output_dir_box], outputs=[setting_result])
    
    gr.Markdown("© @KotaOoka")
    
demo.launch(inbrowser=True)  # 公開リンクが必要なら demo.launch(share=True) としてください

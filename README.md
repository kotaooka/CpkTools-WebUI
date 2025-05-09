# CpkTools-WebUI

**CpkTools WebUI** は、Excelファイルから手軽に工程能力解析を実施できる製造業の品質管理に適したツールです。  
アップロードした Excel データから、解析対象列や上限・下限規格値を指定するだけで、最大値、最小値、平均値、標準偏差、Cp、Cpk、尖度、歪度などの統計量を瞬時に算出し、ヒストグラム、QQプロット、確率密度分布、X-bar/Iチャート、R/MRチャート、s管理図などの多彩なグラフを自動生成します。

## 🚀特徴
- **簡単操作**: Excelファイルをアップロードして、解析対象と規格値を設定するだけ
- **多彩な自動解析**: 統計量の計算とグラフ作成を瞬時に実施
- **柔軟な出力形式**: Excel出力と Web UI プレビューで結果確認が可能
- **豊富なサポート情報**: 初心者でも安心の品質管理解説タブを完備

## 📚目次
- [実行例](#✍実行例)
- [要求環境](#💻要求環境)
- [インストール方法](#🔧インストール方法)
  - [1. Python のインストール](#1-python-のインストール)
  - [2. リリースファイルのダウンロード](#2-リリースファイルのダウンロード)
  - [3. setup.bat を実行して仮想環境の構築とライブラリのインストール](#3-setupbat-を実行して仮想環境の構築とライブラリのインストール)
- [CpkTools-WebUI の起動と使用方法](#▶️cpktools-webui-の起動と使用方法)
- [注意事項](#⚠️注意事項)
- [免責事項](#❗免責事項)
- [Third Party Libraries and Licenses](#📜third-party-libraries-and-licenses)

## ✍実行例
以下は本ツールを実行した際の画面例と出力されるグラフです。
<!-- 1枚目の画像 -->
<p align="center">
  <img src="https://github.com/user-attachments/assets/583016ca-24cb-40ec-9f3e-23e24d2e8fc2" alt="実行例1" width="500">
</p>

<!-- 2枚目以降のグリッドを中央揃えに -->
<div align="center">
  <table>
    <tr>
      <td align="center">
        <img src="https://github.com/user-attachments/assets/e69b42d6-0e36-4831-a9ea-69fc489d0f78" alt="実行例2" width="250">
      </td>
      <td align="center">
        <img src="https://github.com/user-attachments/assets/ce5abb5c-5743-4aa2-811a-1961fbf9d42c" alt="実行例3" width="250">
      </td>
      <td align="center">
        <img src="https://github.com/user-attachments/assets/2964abc7-7be4-4afe-9249-e32083340aff" alt="実行例4" width="250">
      </td>
    </tr>
    <tr>
      <td align="center">
        <img src="https://github.com/user-attachments/assets/c146e538-e294-4c54-8ca2-75958a7e3c7d" alt="X-bar チャート例" width="250">
      </td>
      <td align="center">
        <img src="https://github.com/user-attachments/assets/2714d2ce-a681-4f95-ab46-ba3a901ccc24" alt="R チャート例" width="250">
      </td>
      <td align="center">
        <img src="https://github.com/user-attachments/assets/964ef55c-b964-4d26-aa86-e95288c38409" alt="s 管理図例" width="250">
      </td>
    </tr>
  </table>
</div>


## 💻要求環境

- **OS**: Windows  
- **Python**: Python 3.x  
- **必要な Python ライブラリ**:  pandas, pyarrow, matplotlib, scipy, Pillow, gradio, numpy

## 🔧インストール方法

### 1. Python のインストール

Python をインストールしていない場合は、以下のリンクから最新の Python 3.x をダウンロードしてください。  
※ インストール時に「Add Python to PATH」にチェックを入れることを推奨します。

- [Python 公式ダウンロードページ (Windows)](https://www.python.org/downloads/windows/)

### 2. リリースファイルのダウンロード

本プロジェクトの最新リリースは [GitHub Releases](https://github.com/kotaooka/CpkTools-WebUI/releases) ページからダウンロードできます。  
ダウンロードした ZIP ファイルを展開し、任意のフォルダ（例：`D:\CpkTools-WebUI`）に保存してください。

### 3. setup.bat を実行して仮想環境の構築とライブラリのインストール

解凍したプロジェクトフォルダ内にある `setup.bat` を実行します。  
このバッチファイルは、プロジェクト専用の Python 仮想環境を作成し、必要なライブラリのインストールを自動で行います。
（お使いのネットワーク環境でプロキシ設定が必要な場合は、`setup.bat` に適宜プロキシ設定用の環境変数のコードを追加してください。）

```
setup.batの編集例
set HTTP_PROXY=http://proxy.example.com:8080
python -m venv venv
-以下略-
```
## ▶️CpkTools-WebUI の起動と使用方法

### 1. CpkTools-WebUI.bat を実行してアプリケーションを起動

プロジェクトフォルダ内にある `CpkTools-WebUI.bat` を実行することで、アプリケーションが起動します。  
実行後、ブラウザが自動的に開き、Gradio を用いたインターフェースが表示されます。  
（お使いのネットワーク環境でプロキシ設定が必要な場合は、`CpkTools-WebUI.bat` に適宜プロキシ設定用の環境変数のコードを追加してください。）

```
CpkTools-WebUI.batの編集例
@echo off
set HTTP_PROXY=http://proxy.example.com:8080
rem Check if the virtual environment directory exists
-以下略-
```

### 2. Excel ファイルのアップロードとプレビュー

- 「**Excel ファイル**」アップロードボックスに対象の Excel ファイル（`.xlsx` または `.xls`）を選択します。
- Excelファイルの内容は1列を1群として扱います。 
- アップロード後、ファイルの先頭 5 行がプレビュー表示され、利用可能な列がドロップダウンリストに表示されます。  
- オプションとして「**先頭行をデータとして扱う**」チェックボックスにより、Excel の先頭行も解析対象のデータとして読み込むことが可能です。


### 3. 解析対象の列選択と規格値入力

- 表示されたドロップダウンリストから、解析対象となる列（例：A列、B列など）を複数選択してください。  
- 選択した列に基づいて、自動生成される「**各列の規格値入力**」テーブルに、各列ごとの上限規格値と下限規格値を入力します。  
- すべての列で同一の規格値を使用する場合は、「**すべての列の規格値を同じにする**」チェックボックスを ON にすると、1 列目の規格値が自動的に全列にコピーされます。

### 4. サブグループサイズと標準偏差計算方法の選択

- サブグループサイズは、1 ～ 10 の範囲で設定可能です。  
  - **1** を選択すると、I チャートおよび MR チャート（個々のデータに基づく管理図）が生成されます。  
  - **2 以上** を選択すると、X-bar チャート／R チャートおよび十分なデータがある場合は s 管理図が生成されます。  
- また、標準偏差の計算方法として「サンプル標準偏差」または「母集団標準偏差」から選択できます。

### 5. 解析開始と結果の確認

- 入力内容に問題がなければ、「**解析開始**」ボタンをクリックします。  
- 各列に対して、統計量の計算およびグラフの生成が実行され、解析結果のログが画面上に表示されます。  
- 統計結果は、Web UI 上の結果プレビューおよび Excel プレビュー領域に表示され、生成された Excel ファイルやグラフは `output` フォルダに保存されます。  
- Web UI 上からグラフ（ヒストグラム、QQ プロット、確率密度分布、X-bar／I チャート、R／MR チャート、s 管理図）のプレビューを確認することができ、また各ファイルはダウンロードすることも可能です。  
- ※ s 管理図に関しては、各サブグループで十分なデータ（2点以上）があるグループのみをプロットする仕組みにより、データ点数の不足によるエラーが回避されています。

### 6. 出力フォルダの確認

- 出力されたグラフや Excel ファイルは、`output` フォルダに保存されます。  
- Web UI 上の「Outputフォルダを開く」ボタンを利用すると、Windows のエクスプローラーで出力フォルダが自動的に開きます。

## ⚠️注意事項

- アップロードする Excel ファイルに有効なデータが含まれていること、また解析対象の列に欠損値が極力ないことを確認してください。（欠損値は自動的に除外されますが、連続した欠損データの場合は注意が必要です。）  
- 規格値は正しい数値形式で入力してください。すべての列で同じ規格値を使用する場合は、まず 1 列目に正確な値を入力してからチェックボックスを ON にしてください。  
- サブグループサイズの設定は、解析対象のデータ数に応じた適切な値を選択する必要があります。  

## ❗免責事項

**【本ソフトウェアの提供について】**  
本アプリケーション（以下「本ソフトウェア」といいます）は、現状有姿 ("as is") の状態で提供されています。  
作者および配布者は、本ソフトウェアの動作、正確性、有用性、信頼性について、明示的にも黙示的にも一切の保証を行いません。  
ユーザーは自己責任のもとで本ソフトウェアを利用するものとし、本ソフトウェアの使用または不使用に起因する直接的または間接的な損害、データの消失、システムの不具合等について、作者および配布者は一切責任を負いません。

**【使用環境およびデータの取扱いに関して】**  
本ソフトウェアは、多様な環境での動作を想定して開発されていますが、利用環境や利用方法により予期しない動作やエラーが生じる可能性があります。  
特に、データの転送、解析、及び出力に関しては、ユーザー自身で十分な確認およびバックアップを取得した上でご利用ください。  
利用中に生じたトラブル（例：誤った結果の出力、データの損失等）について、作者および配布者は一切保証または補償をいたしません。

**【ライセンスについて】**  
本ソフトウェアは [MITライセンス](./LICENSE) の下で配布されています。  
MITライセンスに基づき、本ソフトウェアの利用、改変、再配布は自由に行えますが、本免責事項はそのまま適用されます。  
詳細につきましては、同梱の LICENSE ファイルをご参照ください。

**【免責事項の変更について】**  
本免責事項は、予告なく変更される可能性があります。  
最新の免責事項を随時ご確認いただき、その内容に同意の上で本ソフトウェアをご利用ください。  
免責事項に同意いただけない場合は、本ソフトウェアの使用をお控えくださいますようお願いいたします。

## 📜Third Party Libraries and Licenses

このプロジェクトはいくつかのオープンソースライブラリを利用しています。以下の著作権情報にご注意ください：

- **pandas**  
  Licensed under the BSD 3-Clause License.  
  View the full license [here](https://github.com/pandas-dev/pandas/blob/main/LICENSE).

- **pyarrow**  
  Licensed under the Apache License 2.0.  
  View the full license [here](https://github.com/apache/arrow/blob/master/LICENSE.txt).

- **scipy**  
  Licensed under the BSD License.  
  View the full license [here](https://github.com/scipy/scipy/blob/main/LICENSE.txt).

- **matplotlib**  
  Licensed under a BSD-style license.  
  View the full license [here](https://matplotlib.org/stable/users/license.html).

- **pillow**  
  Licensed under the Historical PIL License.  
  View details [here](https://github.com/python-pillow/Pillow/blob/main/LICENSE).

- **gradio**  
  Licensed under the MIT License.  
  View the full license [here](https://github.com/gradio-app/gradio/blob/main/LICENSE).

- **openpyxl**  
  Licensed under the MIT License.  
  View details [here](https://openpyxl.readthedocs.io/en/stable/license.html).

- **numpy**  
  Licensed under the BSD License.  
  View the full license [here](https://github.com/numpy/numpy/blob/main/LICENSE.txt).


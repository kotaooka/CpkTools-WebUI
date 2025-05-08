# CpkTools-WebUI

CpkTools WebUI は、Excel ファイルから工程能力解析を実施できるツールです。  
ユーザーは Excel ファイルをアップロードし、解析対象の列を選択、各列ごとの上限規格値と下限規格値を入力して、  
各種統計値（最大値、最小値、標準偏差、平均値、Cp、Cpk、尖度、歪度）を計算および
グラフ（ヒストグラム、QQプロット、確率密度分布）の生成を行います。  
結果は Excel ファイルに出力され、Web UI 上にプレビュー表示されます。

## 目次
- [実行例](#実行例)
- [要求環境](#要求環境)
- [インストール方法](#インストール方法)
  - [1. Pythonのインストール](#1-pythonのインストール)
  - [2. リリースファイルのダウンロード](#2-リリースファイルのダウンロード)
  - [3. setup.batを実行して仮想環境の構築とライブラリのインストール](#3-setupbatを実行して仮想環境の構築とライブラリのインストール)
- [CpkTools-WebUIの起動と使用方法](#cpktools-webuiの起動と使用方法)
- [注意事項](#注意事項)
- [ライセンス](#ライセンス)


## 実行例
![image](https://github.com/user-attachments/assets/cb18c623-0eba-4da9-872f-88a65000e740)

![image (2)](https://github.com/user-attachments/assets/e69b42d6-0e36-4831-a9ea-69fc489d0f78)

![image (3)](https://github.com/user-attachments/assets/ce5abb5c-5743-4aa2-811a-1961fbf9d42c)

![image (1)](https://github.com/user-attachments/assets/2964abc7-7be4-4afe-9249-e32083340aff)



## 要求環境

- **OS**: Windows  
- **Python**: Python 3.x  
- **必要な Python ライブラリ**: pandas, pyarrow, matplotlib, scipy, Pillow, gradio, numpy

## インストール方法

### 1. Pythonのインストール

Python をインストールしていない場合は、以下のリンクから最新の Python 3.x をダウンロードしてください。  
※ インストール時に「Add Python to PATH」にチェックを入れることを推奨します。

- [Python 公式ダウンロードページ (Windows)](https://www.python.org/downloads/windows/)

### 2. リリースファイルのダウンロード

本プロジェクトの最新リリースは [GitHub Releases](https://github.com/kotaooka/CpkTools-WebUI/releases) ページからダウンロードできます。  
ダウンロードした ZIP ファイルを展開し、任意のフォルダ（例：`D:\CpkTools-WebUI`）に保存してください。

### 3. setup.batを実行して仮想環境の構築とライブラリのインストール

ダウンロードまたは解凍したプロジェクトフォルダ内にある `setup.bat` を実行します。  
このバッチファイルは、プロジェクト専用の Python 仮想環境を作成し、必要なライブラリのインストールを自動で行います。

## CpkTools-WebUIの起動と使用方法

### 1. CpkTools-Webui.batを実行してアプリケーションを起動

同じフォルダ内にある `CpkTools-WebUI.bat` を実行することで、アプリケーションが起動します。  
実行後、ブラウザが自動的に開き、**Gradio** を用いたユーザーインターフェースが表示されます。

**プロキシ設定が必要な環境の場合はCpkTools-WebUI.batファイルに以下のコードを追加してください。**
```
  set HTTP_PROXY=http://あなたのアドレス:ポート番号
  set HTTPS_PROXY=https://あなたのアドレス:ポート番号
  set NO_PROXY=localhost,127.0.0.1
```

### 2. Excel ファイルのアップロード

「**Excelファイル**」アップロードボックスに対象の Excel ファイル（`.xlsx` または `.xls`）を選択してください。  
アップロード後、ファイルの先頭 5 行がプレビュー表示され、利用可能な列がドロップダウンリストに表示されます。

### 3. 解析対象の列選択

表示されたドロップダウンリストから、**解析対象の列**を選択してください（複数選択可）。
![image](https://github.com/user-attachments/assets/114ba46f-20b5-48e4-b101-94b28200d118)


### 4. 各列の規格値入力

選択した列に基づいて、下部に自動生成される「**各列の規格値入力**」テーブルが表示されます。  
テーブルには以下の項目が含まれます：

- **解析対象**: 選択された列名  
- **規格上限値**: 各列ごとの上限規格値  
- **規格下限値**: 各列ごとの下限規格値

すべての列で同一の規格値を使用する場合は、「**すべての列の規格値を同じにする**」チェックボックスを ON にしてください。  
チェックボックスが有効になると、**1 列目の規格値が自動的に全列にコピー**されます。

![image](https://github.com/user-attachments/assets/cbfdd102-3840-4a97-aea2-0e2b37fe10f5)


### 5. 解析開始

入力内容に問題がなければ、「**解析開始**」ボタンをクリックしてください。  
すると、各列に対する統計計算（**最大値、最小値、標準偏差、平均値、Cp、Cpk、尖度、歪度**）と**グラフ生成**が行われ、解析結果のログが画面上に表示されます。

### 6. 結果の確認とダウンロード

- 統計結果は、画面上の**結果プレビュー**と**Excel プレビュー領域**に表示されます。  
- 生成された Excel ファイルとグラフは、`output`フォルダに出力され、**ダウンロードリンク**からも取得できます。  
- 各列の**ヒストグラム**と**QQ プロット**と**確率密度分布**も表示され、解析内容を視覚的に確認できます。

## 注意事項

- アップロードする Excel ファイルに**有効なデータ**が含まれていること、解析対象の列に**欠損値がない**ことを確認してください。
- 規格値入力は**正しい数値形式**で行ってください。すべての列で同じ規格値を使用する場合は、まず**1 列目に正確な値**を入力してからチェックボックスを ON にしてください。

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Third-Party Libraries & Licenses

This project makes use of several open source libraries. Please note the following copyright information:

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
  Licensed under the Historical PIL License (a variant of an open source license similar to BSD).  
  View details [here](https://github.com/python-pillow/Pillow/blob/main/LICENSE).

- **gradio**  
  Licensed under the MIT License.  
  View the full license [here](https://github.com/gradio-app/gradio/blob/main/LICENSE).

- **openpyxl**  
  Licensed under the MIT License.  
  View license details [here](https://openpyxl.readthedocs.io/en/stable/license.html).

- **numpy**  
  Licensed under the BSD License.  
  View the full license [here](https://github.com/numpy/numpy/blob/main/LICENSE.txt).

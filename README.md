## 🔍 Excel検索ツール（Search Tool）

### 🔍 概要

このリポジトリは、複数のExcelファイル（.xlsx）を対象に、指定したキーワードを一括検索し、該当するセルの内容やファイル名、シート名、セル位置などを一覧として出力するPythonツールです。大量データの中から効率的に情報を抽出したい場面に活用できます。

### 💡 主な機能

* 指定フォルダ内のすべてのExcelファイルを走査
* ユーザーが入力したキーワードを検索
* ヒットした情報をCSV形式で保存（ファイル名、シート名、セル番地、内容）

### 🛠️ 使用技術

* Python
* openpyxl / pandas（Excel処理）
* os / pathlib（ファイル操作）

### 🚀 実行方法

```bash
search_tool_Ver06.py
```

### 📦 ファイル構成

```
├── search_tool_Ver06.py # メインスクリプト
├── example.xlsx         # 検索対象のサンプルExcelファイル
├── output.csv           # 検索結果（出力ファイル）
└── README.md            # プロジェクト説明（日本語）
```

### 🧑‍💻 作者

[Masaakiii](https://github.com/Masaakiii)

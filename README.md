 
# AutoTaxHelper 🧾✨

**TaxEase** is a simple yet powerful Excel VBA automation tool that streamlines the Japanese tax filing process for freelancers and small business owners.  
It automatically scans receipts and invoices from designated folders, extracts key data from filenames, and populates structured expense and sales sheets in your Excel workbook.

---

## 📂 Folder Structure

```
TaxEase/
├── source/
│   ├── ExpenseTable.bas           # 経費データ読み込みマクロ
│   ├── SalesTable.bas             # 売上データ読み込みマクロ
│   ├── ProfitStatement.bas        # 損益計算シート更新マクロ
│   └── Main.bas                   # すべてを一括実行するマクロ
├── example/
│   └── TaxEase_Sample.xlsm       # 使用例付きのExcelファイル
├── README.md
└── .gitignore
```

---

## 📄 ファイル命名ルール（※重要）

自動でデータを抽出するため、**PDFファイル名は以下の形式に従ってください：**

```
[年]_[月]_[日]_[勘定科目]_[金額]_[備考]_[税率].pdf
```

### ✅ 例：

```
2025_02_06_雑費_1000_コピー機_10%.pdf
```

| 項目       | 内容例       |
|------------|--------------|
| 年月日     | 2025_02_06   |
| 勘定科目   | 雑費         |
| 金額       | 1000         |
| 備考       | コピー機     |
| 税率       | 10%          |

---

## 🔧 使用手順

### ① 必須：ファイル保存先のフォルダを指定

各マクロ冒頭の以下のコードを**自分の環境に合わせて修正**してください：

```vb
folderPath = "G:\マイドライブ\領収書管理\2025年\経費\"
```

売上は：

```vb
folderPath = "G:\マイドライブ\領収書管理\2025年\請求書\"
```

---

### ② Excel へのマクロインポート手順

1. `.xlsm` ファイルを開く
2. `Alt + F11` → VBAエディタを開く
3. `ファイル → ファイルのインポート` で `.bas` ファイル群を取り込む

---

### ③ ボタンにマクロを割り当てる

1. Excelシートに「挿入 → フォームコントロール → ボタン」を追加
2. ボタンを右クリック → 「マクロの登録」
3. `UpdateAllTablesAndProfit` を選択
4. ボタンに「すべて更新」などとラベルをつけて完成！

---

## 📌 利用可能なマクロ一覧

| マクロ名                     | 説明                         |
|------------------------------|------------------------------|
| `UpdateExpenseTable`         | 経費PDFを読み取り、経費シートを更新 |
| `UpdateSalesTable`           | 売上PDFを読み取り、売上シートを更新 |
| `UpdateProfitStatement`      | 経費と売上の差額から損益を自動計算 |
| `UpdateAllTablesAndProfit`   | 上記すべてを一括で実行（おすすめ） |

---

## 📢 注意点

- ファイル名が正しくないと自動読み込みできません
- 同じファイル名のPDFは二重登録されないようになっています
- 消費税額や税抜金額は自動計算されます

---

## 👨‍💻 作者

**K＆A Tech 神田智弥**  
手動入力にかかる時間を最小限に、正確でストレスフリーな帳簿管理を支援します✨

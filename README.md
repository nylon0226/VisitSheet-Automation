# VisitSheet-Automation

![Excel VBA](https://img.shields.io/badge/-Excel%20VBA-217346?logo=microsoft-excel&logoColor=white)
![Automation](https://img.shields.io/badge/-Automation-blue)
![Portfolio](https://img.shields.io/badge/-Portfolio-black)

訪問パターンブックからご利用者ごとの予定表を自動生成・転記するVBAスクリプトです。


# Excel VBA｜訪問予定表 自動転記システム

## 🧭 概要 / Overview  
このExcel VBAツールは、「サービスチェックシート」から各事業所ごとのFAX送付状を**ダブルクリック操作だけで自動生成**します。  
介護・医療系の業務で、同一書類を複数の事業所へ送る際の事務作業を大幅に削減します。

This Excel VBA tool automatically generates individual FAX cover sheets for each care office based on the master sheet ("サービスチェックシート").  
Just **double-click once** to produce formatted FAX sheets ready for printing or transmission.

---

## ⚙️ 主な機能 / Key Features  

| 機能 | 説明 |
|------|------|
| 🖱️ **ダブルクリックで自動生成** | 「サービスチェックシート」上の任意のセル（A列 or B列）をダブルクリックすると、各事業所のFAX送付状が自動作成されます。 |
| 🧩 **事業所ごとのシート自動生成** | 「FAX原本」テンプレートを複製し、事業所名でシート命名。 |
| 👥 **利用者名の重複除去・整列** | 同一事業所内で重複した利用者名を自動で整理。4名ごとに改行。 |
| 🧾 **FAX送信枚数の自動計算** | `送信枚数 = 利用者数 × 2 + 1` で自動算出。 |
| 🧹 **安全なシート名変換** | 禁止文字（/:\\?*[]）を除去してExcel準拠のシート名を生成。 |

---

## 🎯 使用手順 / How to Use  

1. **Excelマクロを有効化**  
　→ 「セキュリティの警告」が出たら「コンテンツの有効化」をクリック。  

2. **「サービスチェックシート」シートを開く**  
　→ 各事業所名がA列、利用者名がB列に入力されていることを確認。  

3. **任意の事業所名セル（A列）をダブルクリック**  
　→ 自動で以下が行われます：  
　　- 「FAX原本」シートを複製  
　　- 新しいシート名を事業所名に変更  
　　- 該当利用者名リストを転記（重複削除・整形）  
　　- 利用者数からFAX送信枚数を自動計算  

4. **完了メッセージを確認**  
　→ 「シートの作成が完了しました。」が表示されたら、FAX送付状が生成完了です。  

---

## 🧩 トリガー / Trigger  
| トリガータイプ | 対象シート | 対象操作 | 動作内容 |
|----------------|-------------|------------|-----------|
| `Worksheet_BeforeDoubleClick` | `FAX原本` | 任意のセルを**ダブルクリック** | 「サービスチェックシート」を参照し、全事業所のFAX送付状を自動生成 |

📌 **補足：**  
- `FAX原本` シート上でどのセルをダブルクリックしても動作します。  
- 生成対象データは「サービスチェックシート」の A・B 列を参照。  
- 完了後、「シートの作成が完了しました。」と表示されます。
---

## 🧠 処理フロー / Logic Flow  

```plaintext
Double-click on "サービスチェックシート"
        ↓
Extract clicked care office name
        ↓
Filter matching client names (B列)
        ↓
Remove duplicates, format with separators
        ↓
Copy "FAX原本" template sheet
        ↓
Rename sheet to care office name
        ↓
Fill client list and total page count
        ↓
Display completion message

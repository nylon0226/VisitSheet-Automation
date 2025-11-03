# VisitSheet-Automation

![Excel VBA](https://img.shields.io/badge/-Excel%20VBA-217346?logo=microsoft-excel&logoColor=white)
![Automation](https://img.shields.io/badge/-Automation-blue)
![Portfolio](https://img.shields.io/badge/-Portfolio-black)

訪問パターンブックからご利用者ごとの予定表を自動生成・転記するVBAスクリプトです。


# Excel VBA｜訪問予定表 自動転記システム

## 🧩 概要
このプロジェクトは、**訪問予定表（看護）ファイル群**を自動生成・自動転記するための Excel VBA スクリプトです。  
「訪問パターン」ブック内の各ご利用者シート（＝転記元）に記載されたスケジュールを、同フォルダ内の各 `.xlsm` ファイル（＝転記先）へ自動反映します。  
原本シートのコピー・日付判断・月ごとのスケジュール配置を完全自動化します。

---

## ⚙️ 主な処理内容

| 処理 | 内容 |
|------|------|
| ① ご利用者検出 | 「一括作成」「シート検索」以外の各ご利用者シートを順番に処理 |
| ② 転記先判定 | 同フォルダ内で、シート名と同名の `.xlsm` ファイル（例：田中.xlsm）を検索 |
| ③ 原本コピー | 開いた訪問予定表ファイル内の「原本」シートをコピーして新しいシートを作成 |
| ④ 共通情報転記 | A3に「ご利用者 様」、G5に一括作成シートA1の月数を転記 |
| ⑤ シート名変更 | 新シートを「11月」などにリネーム |
| ⑥ スケジュール転記 | 各ご利用者シート（転記元）の予定を、日付列と月を照合して訪問予定表に反映 |
| ⑦ 保存・終了 | 保存して閉じ、次のご利用者へ進む（エラー時もスキップ） |

---

## 📂 フォルダ構成例

---

## 🧭 処理イメージ
| 役割 | 対象 | 内容 |
|------|------|------|
| 転記元 | VisitPattern.xlsm（各ご利用者シート） | ご利用者ごとのスケジュール情報を入力する |
| 転記先 | 各ご利用者の訪問予定表ファイル | 原本を複製し、月間スケジュールを自動反映する |
| 原本 | ■訪問予定表(原本)■.xlsm | 転記先ファイルが参照するテンプレート |

---

## 💡 転記ロジック（自動反映の仕組み）

| サブルーチン | 機能概要 |
|---------------|-----------|
| `ApplyOriginalLogic` | 転記先シート内の日付セルと転記元の月を照合し、対応するブロックにデータを挿入 |
| `secondprg` | 週単位のスケジュール転記。対象月の日付を抽出し、対応行に値を転記 |
| `thirdprg` | 冒頭週および末尾週の追加スケジュールを配置し、月末処理まで対応 |

---

## 💾 GitHubへの保存手順

```bash
git init
git add .
git commit -m "Initial upload of VisitSheet Automation project"
git branch -M main
git remote add origin https://github.com/nylon0226/VisitSheet-Automation.git
git push -u origin main

MIT License
© 2025 nylon0226

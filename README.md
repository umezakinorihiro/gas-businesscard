# gas-businesscard

Google Apps Script（GAS）で、名刺画像（jpg/png）から情報を抽出し、Googleスプレッドシートに1行ずつ追記するツールです。  
スプレッドシートの **サイドバー** から画像を選んで実行します。

---

## Features
- スプレッドシートに「名刺」メニューを追加
- サイドバーから名刺画像をアップロードして解析
- Gemini APIで名刺情報を抽出してシートへ追記
- 出力先は `SHEET_ID`（別スプレッドシート）に対応

---

## Files
- `Code.gs` : メニュー/サイドバー表示、Gemini呼び出し、シート追記
- `Sidebar.html` : サイドバーUI
- `appsscript.json` : GAS設定
- `.clasp.json` : clasp設定（公開するなら除外推奨）

---

## Requirements
- Googleアカウント
- Node.js
- clasp（Google Apps Script CLI）
- Gemini API Key

---

## Setup

## 1) clasp
```bash
npm i -g @google/clasp
clasp login

## 2) Apps Script API をON
以下を開いて **Google Apps Script API** を有効化します。  
https://script.google.com/home/usersettings

## 3) 既存GASをclone
Apps Scriptエディタ → **プロジェクト設定** → **スクリプトID** をコピーして実行：

```bash
clasp clone <SCRIPT_ID>

## Script Properties（重要）
APIキーはコードに直書きせず、Apps Script の **プロジェクト設定 → スクリプトプロパティ** に入れます。

| Key | Required | Description |
|---|---:|---|
| `GEMINI_API_KEY` | ✅ | Gemini APIキー |
| `SHEET_ID` | ✅ | 書き込み先スプレッドシートID（`/d/<ID>/edit` の `<ID>`） |
| `SHEET_NAME` | - | 出力シート名（未設定なら `名刺`） |
| `CARD_FOLDER_ID` | - | 画像をDriveに保存したい場合のフォルダID |

---

## Usage
1. （このGASが紐づく）スプレッドシートを開く  
2. 再読み込みしてメニューに **「名刺」** が出ることを確認  
3. **名刺 → 読み込み（サイドバー）**  
4. 画像を選んで「読み込み」  
5. `SHEET_ID` 側のシートに行が追加されます  

---

## Notes
- 名刺は個人情報です。取り扱いに注意してください。
- 429（Quota exceeded）が出る場合は、Gemini APIの利用枠/課金設定を確認してください。
- 公開リポジトリの場合は `.clasp.json` をコミットしない運用を推奨します。

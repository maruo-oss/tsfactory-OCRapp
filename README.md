# AI-OCR 発注管理システム

Google Apps Script と Gemini AI を使用した、PDF発注書の自動OCR処理・管理システムです。

## 概要

このシステムは、PDF形式の発注書を自動的に読み取り、Google Spreadsheet に発注データを保存します。AI（Gemini 2.5 Pro）により発注書の内容を構造化データとして抽出し、Webインターフェースで確認・編集が可能です。

### 主な機能

- **AI-OCR自動処理**: Gemini API を使ってPDF発注書から情報を自動抽出
- **拠点別管理**: Google Drive のフォルダ構造で複数拠点を管理
- **納品先対応**: 発注書から納品先情報を抽出・管理
- **Webインターフェース**: ブラウザ上で発注内容の閲覧・編集
- **PDFプレビュー**: 原本PDFをリアルタイム表示
- **Excel出力**: 選択した発注書をまとめてExcel形式でエクスポート
- **データ編集**: 発注日、メーカー、商品情報などを手動修正可能

## システム構成

```
Google Drive (PDF保管)
    ↓
Google Apps Script (OCR処理)
    ↓
Google Spreadsheet (データ保存)
    ↓
Web UI (閲覧・編集)
```

## セットアップ

### 1. 必要なもの

- Google アカウント
- Gemini API キー（[Google AI Studio](https://makersuite.google.com/app/apikey) で取得）
- Google Drive のフォルダID × 2（入力用、処理済み用）
- Google Spreadsheet

### 2. Google Apps Script プロジェクトの作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 新規プロジェクトを作成
3. 以下のファイルを作成・配置：
   - `Code.gs`: メインスクリプト
   - `index.html`: UI HTML
   - `appsscript.json`: プロジェクト設定

### 3. スクリプトプロパティの設定

スクリプトエディタで「プロジェクトの設定」→「スクリプト プロパティ」から以下を設定：

| プロパティ名 | 説明 | 例 |
|------------|------|-----|
| `GEMINI_API_KEY` | Gemini APIキー | `AIza...` |
| `ROOT_IN_FOLDER_ID` | 処理対象PDFを配置するフォルダID | `1Abc...` |
| `ROOT_PROCESSED_FOLDER_ID` | 処理済みPDFの移動先フォルダID | `1Def...` |
| `SPREADSHEET_ID` | データ保存先スプレッドシートID | `1Ghi...` |

#### フォルダ構造の例

```
📁 ROOT_IN_FOLDER (処理対象)
  ├─ 📁 東京
  │   ├─ 📄 発注書_001.pdf
  │   └─ 📄 発注書_002.pdf
  └─ 📁 大阪
      └─ 📄 発注書_003.pdf

📁 ROOT_PROCESSED_FOLDER (処理済み)
  ├─ 📁 東京 (自動作成)
  └─ 📁 大阪 (自動作成)
```

### 4. Drive APIの有効化

`appsscript.json` で既に設定済みです：
```json
"enabledAdvancedServices": [
  {
    "userSymbol": "Drive",
    "version": "v3",
    "serviceId": "drive"
  }
]
```

### 5. プロンプトシートの作成

スプレッドシートに「プロンプト」という名前のシートを作成し、A1セルに以下のようなプロンプトを記述：

```
以下のPDF発注書から以下の情報をJSON形式で抽出してください：
- order_date: 発注日（YYYY-MM-DD形式）
- maker_name: メーカー名
- shop_name: 店舗名
- delivery_destination: 納品先（株式会社/有限会社など）
- items: 商品明細の配列
  - product_code: 品番
  - product_name: 商品名
  - quantity: 数量
  - unit_price: 単価
```

### 6. Webアプリとしてデプロイ

1. スクリプトエディタで「デプロイ」→「新しいデプロイ」
2. 種類: Webアプリ
3. 実行ユーザー: 自分
4. アクセスできるユーザー: 組織内（または必要に応じて設定）
5. デプロイを実行してURLを取得

## 使い方

### 発注書の自動処理

1. Google Drive の `ROOT_IN_FOLDER` 配下に拠点名フォルダを作成
2. 各拠点フォルダにPDF発注書を配置
3. Webアプリまたはスクリプトエディタから `processOrders()` を実行
   - Web UI: 「自動取込実行」ボタンをクリック
   - または Apps Script トリガーで定期実行

### Web UIの操作

#### 発注書の閲覧
1. 左サイドバーから発注書を選択
2. 中央ペインで詳細情報を確認
3. 右ペインでPDF原本をプレビュー

#### データの編集
1. 発注書を選択
2. 「編集」ボタンをクリック
3. 必要な項目を修正（発注日、メーカー、商品情報など）
4. 「保存」ボタンで変更を確定

#### Excel出力
1. 出力したい発注書にチェック
2. 「選択した店舗をExcel出力」ボタンをクリック
3. 自動生成されたExcelファイルがダウンロード可能

#### フィルタリング
- ヘッダーの「表示フィルタ」で拠点別に絞り込み可能（全拠点/東京/大阪）

## データ構造

### スプレッドシート「OrderData」シート

| 列名 | 説明 |
|------|------|
| branch_name | 拠点名 |
| file_id | Google Drive ファイルID |
| file_name | ファイル名 |
| status | 処理ステータス |
| order_date | 発注日 |
| maker_name | メーカー名 |
| shop_name | 店舗名 |
| product_code | 品番 |
| product_name | 商品名 |
| unit_price | 単価 |
| quantity | 数量 |
| line_total | 小計（単価×数量） |
| processed_at | 処理日時 |
| delivery_destination | 納品先 |

## カスタマイズ

### AIモデルの変更

`Code.gs` の `MODEL_NAME` を変更：
```javascript
const MODEL_NAME = 'gemini-2.5-pro'; // または他のモデル
```

### プロンプトのカスタマイズ

スプレッドシートの「プロンプト」シート A1セルを編集して、抽出項目や形式を変更できます。

### UIのカスタマイズ

`index.html` のスタイルやレイアウトを編集して、デザインを変更できます。

## トラブルシューティング

### PDFが処理されない場合

1. スクリプトプロパティが正しく設定されているか確認
2. Gemini API キーが有効か確認
3. Drive API が有効化されているか確認
4. フォルダ構造が正しいか確認（拠点フォルダ → PDFファイル）

### デバッグ方法

スクリプトエディタで以下を実行：
```javascript
debugFolderCheck(); // フォルダとファイルの診断
```

### エラーが発生したPDF

処理中にエラーが発生したPDFは、ファイル名に `【ERROR】` プレフィックスが付きます。

## 技術スタック

- **バックエンド**: Google Apps Script (V8 Runtime)
- **AI/OCR**: Google Gemini 2.5 Pro API
- **ストレージ**: Google Drive, Google Spreadsheet
- **フロントエンド**: HTML, CSS, JavaScript（Vanilla）
- **タイムゾーン**: Asia/Tokyo

## ライセンス

このプロジェクトは自由に使用・改変できます。

## 開発者

tsfactory

## 更新履歴

- 初回リリース: AI-OCR機能実装
- 納品先対応版: delivery_destination フィールド追加

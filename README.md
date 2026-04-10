# 進捗管理 PoC

ローカル PC 上で動かすことを前提にした、サーバー不要の Tauri デスクトップアプリです。

- 現在の想定版: v0.1 β版
- フロントエンド: React + Vite
- デスクトップランタイム: Tauri
- 保存先: ローカルまたは共有フォルダ上の Excel ファイル

## できること

- Excel ファイルの保存先パスをアプリ画面から設定
- Excel ファイルを参照ダイアログから選択
- 新しい Excel ファイルの保存先をダイアログから選択
- 進捗一覧の表示と絞り込み
- 進捗の新規登録
- 既存進捗の更新
- 報告用メモの登録と確認
- 直近1週間または2週間の定例報告ドラフト生成
- 会議向け報告の履歴保存と前回差分の自動抽出
- ステータス別の件数サマリ表示

## 前提

- Windows ローカル実行を想定
- 共有フォルダ上の `.xlsx` を直接指定できます
- Excel ファイルをデスクトップ版 Excel で開いたまま更新すると保存競合が起こることがあります
- この PoC は認証なしです
- 同時更新は `Version` 列で衝突検知のみ行います

## セットアップ

```powershell
cd c:\Users\sarah\Documents\Progress
npm.cmd install
```

- サンプルデータ投入スクリプトを使う場合は、初回のみ backend 側の依存関係も入れてください

```powershell
cd c:\Users\sarah\Documents\Progress\backend
npm.cmd install
```

## 開発起動

```powershell
cd c:\Users\sarah\Documents\Progress
npm.cmd run dev
```

- Tauri アプリのウィンドウが起動します

## ビルド

```powershell
cd c:\Users\sarah\Documents\Progress
npm.cmd run build
```

- 生成物は `src-tauri\target\release\bundle\nsis\Progress Tracker PoC_0.1.0-beta.1_x64-setup.exe` です
- 補助的に `src-tauri\target\release\progress-tracker-poc.exe` も生成されますが、配布には installer を使ってください

## β版の起動確認手順

```powershell
cd c:\Users\sarah\Documents\Progress
npm.cmd run build
& ".\src-tauri\target\release\bundle\nsis\Progress Tracker PoC_0.1.0-beta.1_x64-setup.exe"
```

- セットアップ完了後、スタートメニューまたはインストール先のアプリ本体から起動してください
- 初回起動時は設定ファイルが `C:\Users\<ユーザー名>\AppData\Roaming\ProgressTrackerPoc\settings.json` に作成されます
- 初回起動時は Excel ファイルが `C:\Users\<ユーザー名>\Documents\ProgressTrackerPoc\progress.xlsx` に自動生成されます

## サンプルデータ投入

過去 1 か月程度の確認用データを入れる場合は次を実行します。

```powershell
cd c:\Users\sarah\Documents\Progress
node .\scripts\seed-sample.mjs
```

- 既定では `C:\Users\<ユーザー名>\Documents\ProgressTrackerPoc\progress.xlsx` に対して upsert します
- 同じ `RowID` のサンプルは追加ではなく更新されます

## β版の配布手順

現時点では NSIS installer を配布物として扱います。

1. 開発端末で `npm.cmd run build` を実行する
2. `src-tauri\target\release\bundle\nsis\Progress Tracker PoC_0.1.0-beta.1_x64-setup.exe` を配布物として取り出す
3. 必要に応じて初期データ入りの `progress.xlsx` も合わせて配布する
4. 配布先 PC で setup.exe を実行してインストールする
5. インストール後にアプリを起動し、画面上部の設定フォームで利用する Excel パスを確認または変更する

- Excel を同梱しない場合でも、初回起動時に空の Excel が自動生成されます
- サンプル付きで配布したい場合は、事前に sample データを投入した `progress.xlsx` を渡してください
- 現状の `npm.cmd run build` は `tauri build` を実行し、NSIS インストーラを生成します

## Excel ファイルについて

- 初回起動時はユーザーの Documents 配下に `ProgressTrackerPoc/progress.xlsx` が自動生成されます
- 画面上部の設定フォームで任意の `.xlsx` ファイルパスを入力し、Enter またはフォーカス移動で切り替えられます
- 既存ファイルは参照ボタンから選ぶと、その場で切り替わります
- 新規ファイルは新規作成ボタンから保存先を選ぶと、その場で切り替わります
- 共有フォルダの UNC パスも指定できます
- 既存ファイルを使う場合は、アプリが期待するヘッダー列を持つ必要があります

## 現在の構成

- `frontend/`: React UI
- `src-tauri/`: Tauri 本体と Excel 読み書きロジック
- `backend/`: 以前の Web PoC 名残。現行起動では使いません

## 想定ヘッダー

- RowID
- KPI番号
- カテゴリー
- 担当者名
- 登録日
- 更新日
- ステータス
- ランク
- ディールサイズ
- 社外関係者
- 社内関連部署
- 顧客名
- 内容
- NextAction
- 報告メモ
- 更新者
- Version

## 運用上の注意

- 1つの Excel を複数人が同時に編集すると、アプリ側で競合を検知して保存を止める場合があります
- 共有フォルダの権限により保存に失敗する場合があります
- 本番導入前に、共有フォルダ上の実ファイルで 2, 3 人の同時更新テストを行ってください


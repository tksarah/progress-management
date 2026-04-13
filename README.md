
 # 進捗管理 PoC

 ローカルで動作する Tauri ベースのデスクトップアプリ（PoC）。フロントエンドは React + Vite、ネイティブ処理は Rust で実装され、ローカルの Excel (`.xlsx`) をデータ保存先として利用します。

 **主要技術**: React, Vite, Tauri, Rust

 ## 概要

 このリポジトリはデスクトップアプリとしてローカル Excel を保存先に使う PoC です。UI で保存先の Excel を選択または新規作成し、進捗の一覧表示・登録・更新、報告ドラフト生成などを行えます。

 ## 主な機能

 - Excel ファイルの選択・新規作成
 - 進捗の一覧表示・フィルタリング
 - 進捗の作成・更新・報告メモ管理
 - 定例報告（直近1週間/2週間など）ドラフト生成
 - 保存時の簡易衝突検知（`Version` 列ベース）

 ## 動作前提

 - Node.js（推奨: 18 以上）
 - Rust toolchain（`rustup` で stable を推奨）
 - Windows での利用を想定（ただし macOS/Linux でも動作する可能性あり）
 - PowerShell の実行ポリシー等で npm が止まる場合、`npm.cmd` を直接利用してください

 ## リポジトリ構成

 - [frontend](frontend): React/Vite のフロントエンドソース
 - [src-tauri](src-tauri): Tauri/Rust 側のコード（Excel I/O 等）
 - [scripts](scripts): 補助スクリプト（サンプル生成や Excel ジェネレータなど）
 - [capabilities](capabilities), [gen](gen), [icons](icons): 設定・生成物・アイコン類

 ## 開発用セットアップ

 ルートで依存をインストールします。必要に応じて `frontend` 内でも個別に実行してください。

 Windows PowerShell で問題が出る場合は `npm.cmd` を使ってください。

 ```powershell
 # リポジトリルート
 npm install

 # フロントエンドだけ個別にインストールする場合
 cd frontend
 npm install
 cd ..
 ```

 ## 開発起動

 VS Code のタスク `Run Tauri Dev`（またはリポジトリルートで）を実行します:

 ```powershell
 npm run dev
 # PowerShell で問題があれば
 npm.cmd run dev
 ```

 このスクリプトは通常 Vite (frontend) と Tauri (Rust) の開発サーバーを組み合わせて起動します。

 ## ビルド（配布）

 ```powershell
 npm run build
 ```

 ビルド後のバンドルは `src-tauri/target/release/bundle/` 以下に配置されます（プラットフォームと設定による）。

 ## 便利なスクリプト

 - サンプルデータ投入: `node ./scripts/seed-sample.mjs`
 - Excel 生成/ユーティリティ: `scripts/generate-progress-xlsx.mjs`（用途に応じて実行）

 例:

 ```powershell
 node ./scripts/seed-sample.mjs
 node ./scripts/generate-progress-xlsx.mjs
 ```

 ## Excel テンプレート（想定ヘッダーの例）

 アプリは既定のヘッダーを期待します。最小限の想定例:

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

 （既存ファイルを利用する場合はヘッダーが一致することを確認してください）

 ## 運用上の注意

 - 同一ファイルを複数人で同時編集すると競合します。ネットワーク共有での運用は慎重に行ってください。
 - 共有フォルダを使う場合はアクセス権とネットワークの安定性を事前に確認してください。

 ---

 更に反映したい項目（例: インストーラ署名手順、配布ポリシー、具体的な Excel ヘッダ仕様の正確な定義）があれば指示ください。


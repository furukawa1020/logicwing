# デプロイガイド

## Google Apps Scriptへのデプロイ手順

### 前提条件

- Googleアカウント
- Google Apps Script エディタへのアクセス
- （推奨）clasp CLI ツール

---

## 方法1: Webエディタで手動デプロイ（推奨：初心者向け）

### ステップ1: プロジェクト作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 「新しいプロジェクト」をクリック
3. プロジェト名を「ロンリちゃん」に変更

### ステップ2: ファイルの追加

以下の順番でファイルを追加:

#### 2-1. appsscript.json

1. 左サイドバーの「プロジェクトの設定」⚙️ をクリック
2. 「appsscript.json」マニフェスト ファイルをエディタで表示する」をON
3. エディタに戻り、`appsscript.json` を開く
4. このリポジトリの `appsscript.json` の内容をコピペ

#### 2-2. Code.gs (メインスクリプト)

1. 「ファイル」→「新規」→「スクリプト」
2. ファイル名を `Code` に変更（.gs は自動付与）
3. `Code.gs` の内容をコピペ

#### 2-3. 検出エンジン

**Apps Scriptの制限により、.jsファイルは直接アップロードできません。**
以下のコードを `Code.gs` の**末尾に追加**してください:

```javascript
// detector.js の内容をコピー
// ai_analyzer.js の内容をコピー
// randomizer.js の内容をコピー
```

#### 2-4. HTMLファイル

1. 「ファイル」→「新規」→「HTML」
2. 以下の3つのHTMLファイルを作成:
   - `stamp_animation` (stamp_animation.html の内容)
   - `sidebar` (sidebar.html の内容)
   - `settings` (settings.html の内容)

### ステップ3: デプロイ

1. 「デプロイ」→「新しいデプロイ」
2. 「種類の選択」→「ウェブアプリ」
3. 以下を設定:
   - **説明**: ロンリちゃん v1.0.0
   - **次のユーザーとして実行**: 自分
   - **アクセスできるユーザー**: 自分のみ
4. 「デプロイ」をクリック
5. 権限承認画面で承認

### ステップ4: Google Docsで動作確認

1. Google Docsで新しいドキュメントを作成
2. ページを再読み込み（F5）
3. メニューに「🪽 ロンリちゃん」が表示されることを確認

---

## 方法2: clasp CLIで自動デプロイ（推奨：開発者向け）

### 事前準備

```bash
# clasp CLIをインストール
npm install -g @google/clasp

# Googleアカウントでログイン
clasp login
```

### デプロイ手順

```bash
# このリポジトリをクローン
git clone https://github.com/yourusername/logicwing.git
cd logicwing

# 新しいApps Scriptプロジェクトを作成
clasp create --type standalone --title "ロンリちゃん"

# ファイルをプッシュ
clasp push

# ブラウザでプロジェクトを開く
clasp open
```

### .clasp.json の設定

プロジェクトルートに `.clasp.json` を作成:

```json
{
  "scriptId": "YOUR_SCRIPT_ID",
  "rootDir": "./src"
}
```

### ファイル構成の調整

Apps Scriptの制約に合わせてファイルを統合:

```bash
# すべての.jsファイルを1つに結合
cat src/core/*.js > Code.gs

# HTMLファイルはそのまま
cp stamp_animation.html src/
cp sidebar.html src/
cp settings.html src/
```

---

## トラブルシューティング

### エラー: "Script function not found"

**原因**: 関数名が間違っているか、コードが正しくコピーされていない

**解決策**:
1. `Code.gs` の内容を確認
2. `onOpen` 関数が存在するか確認
3. 保存して再度実行

### エラー: "権限が不足しています"

**原因**: Google Docsへのアクセス権限が承認されていない

**解決策**:
1. Apps Scriptエディタで「実行」→「onOpen」を選択
2. 権限承認画面で「権限を確認」
3. Googleアカウントを選択して承認

### メニューに表示されない

**原因**: ドキュメントが更新されていない

**解決策**:
1. Google Docsのページをリロード（F5）
2. それでも表示されない場合は、新しいドキュメントを作成

### アニメーションが動かない

**原因**: HTMLファイルが正しく読み込まれていない

**解決策**:
1. `stamp_animation.html` が正しく追加されているか確認
2. HTMLファイルの拡張子が `.html` になっているか確認
3. Apps Scriptエディタでファイル名を確認（`stamp_animation.html`）

---

## 更新とバージョン管理

### コードを更新する

```bash
# ローカルで変更を加える
# ...

# 変更をプッシュ
clasp push

# バージョンを作成
clasp version "v1.1.0"

# 新しいバージョンをデプロイ
clasp deploy --versionNumber 2 --description "バグ修正"
```

### ロールバック

```bash
# 以前のバージョンに戻す
clasp deployments
clasp undeploy <DEPLOYMENT_ID>
```

---

## セキュリティとベストプラクティス

### APIキーの管理

- **絶対にコードに直接書かない**
- PropertiesServiceを使用して保存
- 設定画面からのみ入力可能にする

### スコープの最小化

`appsscript.json` で必要最小限の権限のみ要求:

```json
{
  "oauthScopes": [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/script.container.ui"
  ]
}
```

### エラーハンドリング

すべての外部API呼び出しに try-catch を追加:

```javascript
try {
  const result = UrlFetchApp.fetch(url);
} catch (e) {
  Logger.log('Error: ' + e.message);
  showMessage('エラーが発生しました');
}
```

---

## パフォーマンス最適化

### キャッシュの活用

```javascript
const cache = CacheService.getUserCache();
cache.put('key', 'value', 21600); // 6時間キャッシュ
```

### バッチ処理

```javascript
// 1文ずつ処理するのではなく、まとめて処理
const paragraphs = body.getParagraphs();
const results = paragraphs.map(p => analyze(p.getText()));
```

---

## サポート

問題が解決しない場合:

1. [GitHub Issues](https://github.com/yourusername/logicwing/issues) で報告
2. [Google Apps Script コミュニティ](https://support.google.com/docs/community)
3. README.md のトラブルシューティングセクションを参照

---

**Happy Deploying! 🪽**

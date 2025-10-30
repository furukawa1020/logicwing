# ロンリちゃん🪽 プロジェクト

## プロジェクト構成

```
logicwing/
├── README.md              メインドキュメント
├── DEPLOY.md              デプロイガイド
├── LICENSE                ライセンスファイル
├── package.json           npm設定
├── appsscript.json        Apps Script設定
├── Code.gs                メインスクリプト（手動用）
├── Code_Complete.gs       統合版スクリプト（完全版）
│
├── src/                   ソースコード
│   ├── core/              コアロジック
│   │   ├── detector.js    論理検出エンジン
│   │   ├── randomizer.js  ランダム検出
│   │   └── ai_analyzer.js AI補助検出
│   │
│   ├── ui/                ユーザーインターフェース
│   │   └── （HTMLファイルはルートに配置）
│   │
│   └── assets/            アセット
│       └── ronri_stamp.svg キャラクターSVG
│
├── stamp_animation.html   スタンプアニメーション
├── sidebar.html           押印履歴サイドバー
└── settings.html          設定画面
```

## 開発環境のセットアップ

### 1. Node.jsのインストール

```bash
# Node.js 14以上が必要
node --version
```

### 2. 依存関係のインストール

```bash
npm install
```

### 3. clasp CLIの設定

```bash
# Googleアカウントでログイン
npm run login

# 新しいプロジェクトを作成
npm run create

# または既存のプロジェクトに接続
# .clasp.json を手動で作成し、scriptIdを設定
```

### 4. コードのプッシュ

```bash
# Apps Scriptにプッシュ
npm run push

# ブラウザでプロジェクトを開く
npm run open
```

## ファイルの役割

### コアファイル

- **appsscript.json**: Apps Scriptの設定ファイル（スコープ、依存関係）
- **Code.gs**: メインエントリーポイント（手動デプロイ用）
- **Code_Complete.gs**: すべてのモジュールを統合した完全版

### UIファイル

- **stamp_animation.html**: スタンプを押すアニメーション画面
- **sidebar.html**: 押印履歴を表示するサイドバー
- **settings.html**: 設定画面

### コアモジュール

- **detector.js**: 論理検出エンジン（正規表現ベース）
- **randomizer.js**: ランダム検出モジュール
- **ai_analyzer.js**: AI補助検出（OpenAI API連携）

## デプロイ方法

### 方法1: Webエディタ（推奨）

1. [DEPLOY.md](DEPLOY.md) の手順に従う
2. すべてのファイルを手動でコピー

### 方法2: clasp CLI

```bash
# プッシュ
npm run push

# デプロイ
npm run deploy
```

## テスト

### 手動テスト

1. Google Docsで新しいドキュメントを作成
2. サンプルテキストを入力:
   ```
   太陽は東から昇る。だから月は西から昇る。
   当然、このことは誰もが知っている。
   つまり、論理は完璧である。
   ```
3. メニューから「🪽 ロンリちゃん」→「論理の飛躍を検出」を実行
4. スタンプが押されることを確認

### サンプルテキスト

**良い例（飛躍あり）**:
```
この商品は人気がある。したがって高品質である。
明らかに、これは最良の選択だ。
```

**悪い例（飛躍なし）**:
```
この商品は高評価を得ている。多くのユーザーが満足している。
品質管理も徹底されている。
```

## トラブルシューティング

### よくある問題

**Q: メニューに表示されない**
A: ページをリロード（F5）してください

**Q: エラー: "Script function not found"**
A: Code.gsの内容が正しくコピーされているか確認

**Q: アニメーションが動かない**
A: HTMLファイル名が正確か確認（stamp_animation.html）

## 貢献

プルリクエストを歓迎します！

### 開発フロー

1. このリポジトリをフォーク
2. 新しいブランチを作成 (`git checkout -b feature/amazing-feature`)
3. 変更をコミット (`git commit -m 'Add amazing feature'`)
4. ブランチにプッシュ (`git push origin feature/amazing-feature`)
5. プルリクエストを作成

## ライセンス

MIT License - 詳細は [LICENSE](LICENSE) を参照

## お問い合わせ

- Issues: [GitHub Issues](https://github.com/yourusername/logicwing/issues)
- Email: your.email@example.com

---

Made with ❤️ and 🪽

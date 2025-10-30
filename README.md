# 論理の飛躍スタンプ拡張機能（ロンリちゃん）

論理の飛躍を検出してスタンプを押す可愛い拡張機能です。

##  プロジェクト構成

このリポジトリには2つのバージョンが含まれています：

### 1 Microsoft Word版（メイン） - `word-addin/`
- **現在のメインバージョン**
- Office.js を使用したWord Add-in
- TypeScript + Webpack
- 詳細: [word-addin/README.md](word-addin/README.md)

### 2 Google Docs版（レガシー） - `google-docs-version/`
- 初期バージョン（参考用）
- Google Apps Script実装
- 完全動作版

##  クイックスタート（Word版）

```powershell
cd word-addin
npm install
npm start
```

Wordが自動的に起動し、アドインがサイドロードされます！

##  主な機能

###  論理の飛躍検出
- パターンマッチング：論理接続詞を使用した飛躍を検出
- 文体分析：整いすぎた文体やAI生成っぽい文章を検出
- ランダム検出：直感的な検出機能
- コンテキスト評価：前後の文脈からスコアリング

###  ロンリちゃんアニメーション
- 印面と翼を使った可愛いアニメーション
- ふわふわ浮遊＋羽ばたきエフェクト
- 検出時の効果音

###  その他機能
- 押印履歴の保存と表示
- 統計情報（総押印数、平均スコア等）
- カスタマイズ可能な設定

##  ディレクトリ構造

```
logicwing/
 word-addin/              # Word Add-in版（メイン）
    src/
       detector/       # 検出エンジン
       taskpane/       # UI
    assets/             # 画像（印面.png、論理の翼.png）
    manifest.xml        # アドイン設定
 google-docs-version/    # Google Docs版（参考）
    Code_Complete.gs
    stamp_animation.html
    ...
 src/assets/             # 共有アセット
     印面.png
     論理の翼.png
```

##  アセット

ロンリちゃんの画像アセット：
- **印面.png**: ロンリちゃん本体
- **論理の翼.png**: 翼パーツ

##  ドキュメント

- [Word版開発ガイド](WORD_ADDIN_GUIDE.md)
- [Word版README](word-addin/README.md)

##  技術スタック

### Word版
- TypeScript
- Office.js
- Webpack
- CSS3 Animations

### Google Docs版
- Google Apps Script
- HTML/CSS/JavaScript

##  使い方

1. Wordで文書を開く
2. アドインのタスクペインを表示
3. 「論理の飛躍を検出」ボタンをクリック
4. ロンリちゃんが論理の飛躍を見つけてくれます！

##  貢献

プルリクエストを歓迎します！

##  ライセンス

MIT License

##  作者

furukawa1020

---

**ロンリちゃんと一緒に、論理的な文章を書きましょう！**

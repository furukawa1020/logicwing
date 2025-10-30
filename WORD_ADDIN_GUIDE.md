# 🪽 ロンリちゃん Word用アドイン 構築ガイド

## 📁 プロジェクト構造

```
.vscode/
├── logicwing/              ← 現在のフォルダ（Google Apps Script版）
│   ├── Code.gs
│   ├── stamp_animation.html
│   ├── sidebar.html
│   ├── settings.html
│   └── src/assets/
│       ├── 印面.png
│       └── 論理の翼.png
│
└── word-addin/
    └── ロンリちゃん/        ← Word用アドインプロジェクト
        ├── src/
        ├── manifest.xml
        └── package.json
```

## ✅ 完了した作業

1. ✅ Node.js v18.19.0 確認済み
2. ✅ npm v10.2.3 確認済み
3. ✅ Yeoman & generator-office インストール済み
4. ✅ Word用アドインプロジェクト作成完了
5. ✅ npm install 完了（警告あるが動作OK）

## 🚀 次のステップ

### 1. Word用に検出エンジンを移植

**ファイルをコピー：**
```powershell
# logicwingフォルダから実行
Copy-Item src/assets/*.png ../word-addin/ロンリちゃん/assets/
```

**検出エンジンを作成：**
- `../word-addin/ロンリちゃん/src/detector.ts` を作成
- Google Apps Script版の`LogicLeapDetector`クラスをTypeScriptに変換

### 2. Office.js APIを使用してWord連携

**`taskpane.ts`に実装：**
```typescript
// Word文書からテキストを取得
await Word.run(async (context) => {
  const body = context.document.body;
  body.load("text");
  await context.sync();
  
  const text = body.text;
  // 検出エンジン実行
  const detector = new LogicLeapDetector();
  const leaps = detector.detect(text);
  
  // スタンプアニメーション表示
  showStampAnimation(leaps);
});
```

### 3. アニメーションUIを実装

**`taskpane.html`をカスタマイズ：**
- 印面.pngと論理の翼.pngを表示
- CSS Animationsで羽根パタパタ
- 浮遊エフェクト追加

### 4. ローカルでテスト実行

```powershell
cd ../word-addin/ロンリちゃん
npm start
```

これでWordが起動し、アドインがサイドロードされます！

### 5. デプロイ準備

**AppSource公開（本格版）：**
- Microsoft Partner Centerでアカウント作成
- マニフェストの詳細設定
- アイコン・スクリーンショット準備
- 審査申請

**組織内配布（簡易版）：**
- SharePointまたはOneDriveでマニフェストを共有
- ユーザーが手動でインストール

## 🎨 カスタマイズポイント

### 検出パターンの調整
`src/detector.ts`で日本語特有のパターンを追加：
```typescript
const patterns = [
  { regex: /だから|したがって/, weight: 0.7 },
  { regex: /当然|明らか/, weight: 0.8 },
  // 追加のパターン
];
```

### アニメーションの調整
`taskpane.html`のCSSで：
```css
@keyframes wingFlap {
  0%, 100% { transform: rotateZ(0deg); }
  25% { transform: rotateZ(-15deg); }
  75% { transform: rotateZ(15deg); }
}
```

## 📝 コマンドリファレンス

```powershell
# 開発サーバー起動
cd ../word-addin/ロンリちゃん
npm start

# ビルド
npm run build

# Wordにサイドロード（デバッグ）
npm run start:desktop

# Web版Wordで実行
npm run start:web
```

## 🔧 トラブルシューティング

### Node.jsバージョン警告
- Node v18でも動作します（警告は無視してOK）
- Node v20+に更新すると警告が消えます

### Wordが起動しない
```powershell
# Office開発者モードを有効化
npx office-addin-dev-settings appcontainer <manifest-path>
```

### SSL証明書エラー
```powershell
# 開発用証明書をインストール
npx office-addin-dev-certs install
```

## 🎯 完成イメージ

1. **Word文書を開く**
2. **「ホーム」タブに「🪽 ロンリちゃん」ボタンが表示**
3. **クリックすると右側にタスクペイン表示**
4. **「論理の飛躍を検出」ボタンをクリック**
5. **論理ちゃんがアニメーションで登場！**
6. **検出結果がリスト表示**

## 📚 参考リンク

- [Office Add-ins Documentation](https://learn.microsoft.com/office/dev/add-ins/)
- [Word JavaScript API](https://learn.microsoft.com/javascript/api/word)
- [Office UI Fabric](https://developer.microsoft.com/fluentui)

---

**次回から作業を再開する場合：**
```powershell
cd C:\Users\wakuw\OneDrive\画像\デスクトップ\.vscode\word-addin\ロンリちゃん
npm start
```

🎉 Word用ロンリちゃんアドインの開発、頑張ってください！

/**
 * 論理の飛躍スタンプ拡張機能（ロンリちゃん🪽）
 * メインエントリーポイント
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('🪽 ロンリちゃん')
    .addItem('論理の飛躍を検出', 'detectLogicalLeaps')
    .addItem('選択範囲に手動で押印', 'manualStamp')
    .addSeparator()
    .addItem('押印履歴を表示', 'showHistory')
    .addItem('設定', 'showSettings')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

/**
 * 論理の飛躍を自動検出して押印
 */
function detectLogicalLeaps() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const text = body.getText();
  
  // 検出エンジン起動
  const detector = new LogicLeapDetector();
  const leaps = detector.detect(text);
  
  if (leaps.length === 0) {
    showMessage('論理の飛躍は検出されませんでした！✨', '優秀です');
    return;
  }
  
  // サイドバーを表示してアニメーション開始
  showStampAnimation(leaps);
}

/**
 * 選択範囲に手動で押印
 */
function manualStamp() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (!selection) {
    showMessage('テキストを選択してください', 'エラー');
    return;
  }
  
  const elements = selection.getRangeElements();
  if (elements.length === 0) {
    showMessage('有効なテキストが選択されていません', 'エラー');
    return;
  }
  
  const element = elements[0].getElement();
  const text = element.asText().getText();
  
  // 手動押印を記録
  const leap = {
    text: text,
    score: 100, // 手動押印は100%
    type: 'manual',
    position: 0
  };
  
  showStampAnimation([leap]);
}

/**
 * 押印履歴を表示
 */
function showHistory() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('🪽 押印履歴')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * 設定画面を表示
 */
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('settings')
    .setWidth(400)
    .setHeight(300);
  DocumentApp.getUi().showModalDialog(html, '⚙️ ロンリちゃん設定');
}

/**
 * スタンプアニメーションを表示
 */
function showStampAnimation(leaps) {
  const template = HtmlService.createTemplateFromFile('stamp_animation');
  template.leaps = JSON.stringify(leaps);
  
  const html = template.evaluate()
    .setWidth(500)
    .setHeight(400);
  
  DocumentApp.getUi().showModalDialog(html, '');
}

/**
 * メッセージを表示
 */
function showMessage(message, title) {
  DocumentApp.getUi().alert(title || 'お知らせ', message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * 押印データを保存
 */
function saveStampData(leapData) {
  const userProperties = PropertiesService.getUserProperties();
  const historyKey = 'ronri_history';
  
  let history = [];
  const stored = userProperties.getProperty(historyKey);
  if (stored) {
    try {
      history = JSON.parse(stored);
    } catch (e) {
      history = [];
    }
  }
  
  // 新しい押印データを追加
  const stampEntry = {
    text: leapData.text,
    score: leapData.score,
    type: leapData.type,
    timestamp: new Date().toISOString()
  };
  
  history.unshift(stampEntry); // 最新を先頭に
  
  // 最大100件まで保存
  if (history.length > 100) {
    history = history.slice(0, 100);
  }
  
  userProperties.setProperty(historyKey, JSON.stringify(history));
  return true;
}

/**
 * 押印履歴を取得
 */
function getStampHistory() {
  const userProperties = PropertiesService.getUserProperties();
  const historyKey = 'ronri_history';
  const stored = userProperties.getProperty(historyKey);
  
  if (!stored) {
    return [];
  }
  
  try {
    return JSON.parse(stored);
  } catch (e) {
    return [];
  }
}

/**
 * 設定を保存
 */
function saveSettings(settings) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('ronri_settings', JSON.stringify(settings));
  return true;
}

/**
 * 設定を取得
 */
function getSettings() {
  const userProperties = PropertiesService.getUserProperties();
  const stored = userProperties.getProperty('ronri_settings');
  
  const defaultSettings = {
    theme: 'logical_leap',
    randomRate: 5,
    enableSound: true,
    enableAI: false,
    aiApiKey: ''
  };
  
  if (!stored) {
    return defaultSettings;
  }
  
  try {
    return Object.assign(defaultSettings, JSON.parse(stored));
  } catch (e) {
    return defaultSettings;
  }
}

/**
 * 論理の飛躍検出エンジン
 */
class LogicLeapDetector {
  constructor() {
    this.patterns = [
      // 因果関係の飛躍
      { regex: /だから|ゆえに|したがって|よって(?!.*なぜなら|.*理由)/g, type: 'causal', weight: 0.7 },
      // 一般化の飛躍
      { regex: /すべて|全て|みんな|誰でも|必ず|絶対/g, type: 'generalization', weight: 0.8 },
      // 前提の飛躍
      { regex: /当然|明らか|言うまでもなく|疑いようもなく/g, type: 'assumption', weight: 0.6 },
      // 二分法の誤謬
      { regex: /AかBか|～しかない|～以外にない/g, type: 'false_dichotomy', weight: 0.7 },
      // 循環論法
      { regex: /なぜなら.*だからだ|.*ので.*である/g, type: 'circular', weight: 0.5 }
    ];
  }
  
  detect(text) {
    const leaps = [];
    const settings = getSettings();
    
    // パターンマッチング検出
    for (const pattern of this.patterns) {
      const matches = text.matchAll(pattern.regex);
      for (const match of matches) {
        leaps.push({
          text: this.extractContext(text, match.index, 40),
          score: Math.round(pattern.weight * 100),
          type: pattern.type,
          position: match.index
        });
      }
    }
    
    // ランダム検出（設定された確率で）
    if (settings.randomRate > 0 && Math.random() * 100 < settings.randomRate) {
      const randomPos = Math.floor(Math.random() * text.length);
      leaps.push({
        text: this.extractContext(text, randomPos, 40),
        score: Math.round(Math.random() * 50 + 50),
        type: 'random',
        position: randomPos
      });
    }
    
    return leaps;
  }
  
  extractContext(text, position, length) {
    const start = Math.max(0, position - length / 2);
    const end = Math.min(text.length, position + length / 2);
    let context = text.substring(start, end);
    
    if (start > 0) context = '...' + context;
    if (end < text.length) context = context + '...';
    
    return context;
  }
}

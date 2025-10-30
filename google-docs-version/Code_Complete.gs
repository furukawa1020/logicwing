/**
 * 論理の飛躍スタンプ拡張機能（ロンリちゃん🪽）
 * 統合版 Code.gs - すべての機能を含む完全版
 * 
 * このファイルはGoogle Apps Scriptで直接使用できるように
 * すべてのモジュールを統合しています
 */

// ============================================================
// メイン機能
// ============================================================

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

function detectLogicalLeaps() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const text = body.getText();
  
  if (!text || text.trim().length === 0) {
    showMessage('文章が入力されていません', 'エラー');
    return;
  }
  
  const detector = new LogicLeapDetector();
  const leaps = detector.detect(text);
  
  if (leaps.length === 0) {
    showMessage('論理の飛躍は検出されませんでした！✨\n文章は論理的に整っています。', '優秀です');
    return;
  }
  
  showStampAnimation(leaps);
}

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
  
  const leap = {
    text: text.substring(0, 100),
    score: 100,
    type: 'manual',
    position: 0,
    reason: '手動押印'
  };
  
  showStampAnimation([leap]);
}

function showHistory() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('🪽 押印履歴')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}

function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('settings')
    .setWidth(400)
    .setHeight(600);
  DocumentApp.getUi().showModalDialog(html, '⚙️ ロンリちゃん設定');
}

function showStampAnimation(leaps) {
  const template = HtmlService.createTemplateFromFile('stamp_animation');
  template.leaps = JSON.stringify(leaps);
  
  const html = template.evaluate()
    .setWidth(500)
    .setHeight(400);
  
  DocumentApp.getUi().showModalDialog(html, '');
}

function showMessage(message, title) {
  DocumentApp.getUi().alert(title || 'お知らせ', message, DocumentApp.getUi().ButtonSet.OK);
}

// ============================================================
// データ管理
// ============================================================

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
  
  const stampEntry = {
    text: leapData.text,
    score: leapData.score,
    type: leapData.type,
    reason: leapData.reason,
    timestamp: new Date().toISOString()
  };
  
  history.unshift(stampEntry);
  
  if (history.length > 100) {
    history = history.slice(0, 100);
  }
  
  userProperties.setProperty(historyKey, JSON.stringify(history));
  return true;
}

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

function saveSettings(settings) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('ronri_settings', JSON.stringify(settings));
  return true;
}

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

// ============================================================
// 論理検出エンジン (LogicLeapDetector)
// ============================================================

class LogicLeapDetector {
  constructor() {
    this.logicalConnectors = [
      'だから', 'したがって', 'ゆえに', 'よって',
      'つまり', 'すなわち', '要するに',
      'このように', 'このため', 'そのため',
      'それゆえ', '故に', 'ですから', 'なので'
    ];
    
    this.leapPatterns = [
      { pattern: /(.{20,})[。、]\s*(だから|したがって|ゆえに|よって)(.{20,})/g, weight: 0.7 },
      { pattern: /(当然|明らか|言うまでもなく|周知の通り|疑いようもなく)(.{10,})/g, weight: 0.8 },
      { pattern: /(.{30,})[。]\s*(つまり|すなわち|要するに)(.{10,})/g, weight: 0.6 },
      { pattern: /(簡単|容易|一目瞭然|明白)(.{10,})/g, weight: 0.65 },
      { pattern: /(.{40,})[。]\s*(このように|このため|そのため)(.{15,})/g, weight: 0.65 }
    ];
    
    this.tooPolishedPatterns = [
      /重要です。(.{5,20})重要です/g,
      /注意が必要です。(.{5,20})注意が必要です/g,
      /考えられます。(.{5,20})考えられます/g,
      /(まず|次に|最後に)(.{20,})(まず|次に|最後に)/g,
      /です。(.{5,15})です。(.{5,15})です/g
    ];
  }
  
  detect(text) {
    const leaps = [];
    const settings = getSettings();
    
    leaps.push(...this.detectByPatterns(text));
    leaps.push(...this.detectTooPolished(text));
    
    if (settings.randomRate > 0) {
      const randomDetector = new RandomDetector(settings.randomRate);
      leaps.push(...randomDetector.detect(text));
    }
    
    return this.deduplicateAndScore(leaps);
  }
  
  detectByPatterns(text) {
    const leaps = [];
    
    for (const { pattern, weight } of this.leapPatterns) {
      const regex = new RegExp(pattern.source, pattern.flags);
      let match;
      
      while ((match = regex.exec(text)) !== null) {
        const matchedText = match[0];
        const position = match.index;
        
        const contextScore = this.calculateContextScore(text, position, matchedText.length);
        const finalScore = Math.min(100, Math.round(weight * contextScore * 100));
        
        if (finalScore >= 50) {
          leaps.push({
            text: matchedText.substring(0, 100) + (matchedText.length > 100 ? '...' : ''),
            score: finalScore,
            type: 'pattern',
            position: position,
            reason: this.getReasonForPattern(matchedText)
          });
        }
      }
    }
    
    return leaps;
  }
  
  detectTooPolished(text) {
    const leaps = [];
    
    for (const pattern of this.tooPolishedPatterns) {
      const regex = new RegExp(pattern.source, pattern.flags);
      let match;
      
      while ((match = regex.exec(text)) !== null) {
        leaps.push({
          text: match[0].substring(0, 100),
          score: 75,
          type: 'polished',
          position: match.index,
          reason: '整いすぎた文体です🤖'
        });
      }
    }
    
    return leaps;
  }
  
  calculateContextScore(text, position, length) {
    const before = text.substring(Math.max(0, position - 100), position);
    const after = text.substring(position + length, Math.min(text.length, position + length + 100));
    
    let score = 1.0;
    
    if (before.length < 20) score *= 0.8;
    if (after.length < 20) score *= 0.8;
    
    const connectorCount = this.countConnectors(before + after);
    if (connectorCount >= 2) score *= 1.2;
    
    if (/です。(.{5,20})です。/.test(before + after)) {
      score *= 1.3;
    }
    
    return Math.min(1.5, score);
  }
  
  countConnectors(text) {
    let count = 0;
    for (const connector of this.logicalConnectors) {
      if (text.includes(connector)) count++;
    }
    return count;
  }
  
  getReasonForPattern(text) {
    if (text.includes('だから') || text.includes('したがって')) {
      return '論理の飛躍を検出しました';
    }
    if (text.includes('つまり') || text.includes('要するに')) {
      return '要約が唐突かもしれません';
    }
    if (text.includes('当然') || text.includes('明らか')) {
      return '前提が不明確な可能性があります';
    }
    return '論理の飛躍の可能性があります';
  }
  
  deduplicateAndScore(leaps) {
    const merged = [];
    const sorted = leaps.sort((a, b) => a.position - b.position);
    
    for (let i = 0; i < sorted.length; i++) {
      const current = sorted[i];
      
      const isDuplicate = merged.some(m => 
        Math.abs(m.position - current.position) < 50
      );
      
      if (!isDuplicate) {
        merged.push(current);
      }
    }
    
    return merged.slice(0, 10);
  }
}

// ============================================================
// ランダム検出器 (RandomDetector)
// ============================================================

class RandomDetector {
  constructor(rate) {
    this.rate = rate || 5;
    
    this.whimsicalMessages = [
      'なんとなく気になりました🪽',
      'ここ、ちょっと飛んでる気がする...',
      '勘です✨',
      'ピンときました',
      '羽がざわつきました',
      '直感的に引っかかりました',
      'なぜかここに反応しちゃう',
      'ふとした違和感を感じました',
      'スルーできませんでした',
      '気のせいかもしれませんが...'
    ];
  }
  
  detect(text) {
    if (this.rate <= 0) {
      return [];
    }
    
    const sentences = this.extractSentences(text);
    const detections = [];
    
    for (const sentence of sentences) {
      if (Math.random() * 100 < this.rate) {
        const detection = this.createRandomDetection(sentence);
        detections.push(detection);
      }
    }
    
    return detections;
  }
  
  extractSentences(text) {
    const sentences = [];
    const parts = text.split(/[。!?]/);
    
    let position = 0;
    for (const part of parts) {
      const trimmed = part.trim();
      if (trimmed.length >= 10) {
        sentences.push({
          text: trimmed,
          position: text.indexOf(part, position)
        });
        position += part.length + 1;
      }
    }
    
    return sentences;
  }
  
  createRandomDetection(sentence) {
    const message = this.getRandomMessage();
    const score = Math.round(30 + Math.random() * 30);
    
    return {
      text: sentence.text.substring(0, 80) + (sentence.text.length > 80 ? '...' : ''),
      score: score,
      type: 'random',
      position: sentence.position,
      reason: message
    };
  }
  
  getRandomMessage() {
    const index = Math.floor(Math.random() * this.whimsicalMessages.length);
    return this.whimsicalMessages[index];
  }
}

// ============================================================
// AI検出器 (AIAnalyzer) - オプション
// ============================================================

class AIAnalyzer {
  constructor(apiKey) {
    this.apiKey = apiKey;
    this.apiEndpoint = 'https://api.openai.com/v1/chat/completions';
    this.model = 'gpt-3.5-turbo';
  }
  
  analyzeLogicalLeaps(text) {
    if (!this.apiKey) {
      throw new Error('API Keyが設定されていません');
    }
    
    const prompt = this.buildPrompt(text);
    const response = this.callAPI(prompt);
    
    return this.parseResponse(response);
  }
  
  buildPrompt(text) {
    return `以下の文章から論理の飛躍を最大5箇所検出してください。JSON形式で返してください。

文章:
${text}

出力形式:
{"leaps":[{"text":"抜粋","position":0,"score":70,"reason":"理由"}]}`;
  }
  
  callAPI(prompt) {
    try {
      const response = UrlFetchApp.fetch(this.apiEndpoint, {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': `Bearer ${this.apiKey}`
        },
        payload: JSON.stringify({
          model: this.model,
          messages: [
            { role: 'system', content: '論理分析の専門家です。' },
            { role: 'user', content: prompt }
          ],
          temperature: 0.3,
          max_tokens: 1000
        }),
        muteHttpExceptions: true
      });
      
      const result = JSON.parse(response.getContentText());
      
      if (result.error) {
        throw new Error(result.error.message);
      }
      
      return result.choices[0].message.content;
      
    } catch (e) {
      Logger.log(`API呼び出しエラー: ${e.message}`);
      throw e;
    }
  }
  
  parseResponse(responseText) {
    try {
      const jsonMatch = responseText.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        return [];
      }
      
      const data = JSON.parse(jsonMatch[0]);
      
      if (!data.leaps || !Array.isArray(data.leaps)) {
        return [];
      }
      
      return data.leaps.map(leap => ({
        text: leap.text || '',
        score: Math.min(100, Math.max(0, leap.score || 50)),
        type: 'ai',
        position: leap.position || 0,
        reason: leap.reason || 'AI検出'
      }));
      
    } catch (e) {
      Logger.log(`レスポンスのパースエラー: ${e.message}`);
      return [];
    }
  }
}

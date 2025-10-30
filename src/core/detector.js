/**
 * 論理の飛躍検出エンジン (detector.js)
 * 軽量版：正規表現＋接続詞解析
 */

class LogicLeapDetector {
  constructor() {
    // 論理接続詞パターン
    this.logicalConnectors = [
      'だから', 'したがって', 'ゆえに', 'よって',
      'つまり', 'すなわち', '要するに',
      'このように', 'このため', 'そのため',
      'それゆえ', '故に', 'ですから'
    ];
    
    // 飛躍しやすいパターン
    this.leapPatterns = [
      { pattern: /(.{20,})[。、]\s*(だから|したがって|ゆえに)(.{20,})/g, weight: 0.7 },
      { pattern: /(当然|明らか|言うまでもなく|周知の通り)(.{10,})/g, weight: 0.8 },
      { pattern: /(.{30,})[。]\s*(つまり|すなわち|要するに)(.{10,})/g, weight: 0.6 },
      { pattern: /(簡単|容易|一目瞭然|疑いようもなく)(.{10,})/g, weight: 0.7 },
      { pattern: /(.{40,})[。]\s*(このように|このため|そのため)(.{15,})/g, weight: 0.65 }
    ];
    
    // 整いすぎパターン（AI特有）
    this.tooPolishedPatterns = [
      /重要です。(.{5,20})重要です/g,
      /注意が必要です。(.{5,20})注意が必要です/g,
      /考えられます。(.{5,20})考えられます/g,
      /(まず|次に|最後に)(.{20,})(まず|次に|最後に)/g
    ];
  }
  
  /**
   * テキスト全体から論理の飛躍を検出
   */
  detect(text) {
    const leaps = [];
    const settings = this.getSettings();
    
    // 1. パターンマッチング検出
    leaps.push(...this.detectByPatterns(text));
    
    // 2. 整いすぎ検出
    leaps.push(...this.detectTooPolished(text));
    
    // 3. ランダム検出（遊び心）
    if (settings.randomRate > 0) {
      leaps.push(...this.randomDetect(text, settings.randomRate));
    }
    
    // 4. AI補助検出（オプション）
    if (settings.enableAI && settings.aiApiKey) {
      // AI検出は非同期なので、実装は別途
      // leaps.push(...this.detectByAI(text, settings.aiApiKey));
    }
    
    // 重複除去とスコアリング
    return this.deduplicateAndScore(leaps);
  }
  
  /**
   * パターンマッチングによる検出
   */
  detectByPatterns(text) {
    const leaps = [];
    
    for (const { pattern, weight } of this.leapPatterns) {
      let match;
      const regex = new RegExp(pattern.source, pattern.flags);
      
      while ((match = regex.exec(text)) !== null) {
        const matchedText = match[0];
        const position = match.index;
        
        // 文脈スコアを計算
        const contextScore = this.calculateContextScore(text, position, matchedText.length);
        const finalScore = Math.min(100, Math.round(weight * contextScore * 100));
        
        if (finalScore >= 50) { // 50%以上のみ検出
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
  
  /**
   * 整いすぎ文体の検出
   */
  detectTooPolished(text) {
    const leaps = [];
    
    for (const pattern of this.tooPolishedPatterns) {
      let match;
      const regex = new RegExp(pattern.source, pattern.flags);
      
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
  
  /**
   * ランダム検出（気まぐれ）
   */
  randomDetect(text, rate) {
    const leaps = [];
    const sentences = text.split(/[。!?]/);
    
    for (let i = 0; i < sentences.length; i++) {
      const sentence = sentences[i].trim();
      if (sentence.length < 10) continue;
      
      if (Math.random() * 100 < rate) {
        const messages = [
          'なんとなく気になりました🪽',
          'ここ、ちょっと飛んでる気がする...',
          '勘です✨',
          'ピンときました',
          '羽がざわつきました'
        ];
        
        leaps.push({
          text: sentence.substring(0, 80),
          score: Math.round(30 + Math.random() * 30), // 30-60%
          type: 'random',
          position: text.indexOf(sentence),
          reason: messages[Math.floor(Math.random() * messages.length)]
        });
      }
    }
    
    return leaps;
  }
  
  /**
   * 文脈スコアを計算
   */
  calculateContextScore(text, position, length) {
    const before = text.substring(Math.max(0, position - 100), position);
    const after = text.substring(position + length, Math.min(text.length, position + length + 100));
    
    let score = 1.0;
    
    // 前後の文が短すぎる場合はスコアを下げる
    if (before.length < 20) score *= 0.8;
    if (after.length < 20) score *= 0.8;
    
    // 接続詞の密度が高い場合はスコアを上げる
    const connectorCount = this.countConnectors(before + after);
    if (connectorCount >= 2) score *= 1.2;
    
    // 同じ語尾が続く場合（AI特有）
    if (/です。(.{5,20})です。/.test(before + after)) {
      score *= 1.3;
    }
    
    return Math.min(1.5, score);
  }
  
  /**
   * 接続詞の出現回数をカウント
   */
  countConnectors(text) {
    let count = 0;
    for (const connector of this.logicalConnectors) {
      if (text.includes(connector)) count++;
    }
    return count;
  }
  
  /**
   * 検出理由を取得
   */
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
  
  /**
   * 重複除去とスコアリング
   */
  deduplicateAndScore(leaps) {
    // 位置が近い検出結果を統合
    const merged = [];
    const sorted = leaps.sort((a, b) => a.position - b.position);
    
    for (let i = 0; i < sorted.length; i++) {
      const current = sorted[i];
      
      // 既存の検出と近すぎる場合はスキップ
      const isDuplicate = merged.some(m => 
        Math.abs(m.position - current.position) < 50
      );
      
      if (!isDuplicate) {
        merged.push(current);
      }
    }
    
    return merged.slice(0, 10); // 最大10件まで
  }
  
  /**
   * 設定を取得（Apps Scriptから）
   */
  getSettings() {
    try {
      return getSettings();
    } catch (e) {
      return {
        randomRate: 5,
        enableAI: false,
        aiApiKey: ''
      };
    }
  }
}

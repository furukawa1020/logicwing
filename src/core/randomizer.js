/**
 * ランダム検出モジュール (randomizer.js)
 * 気まぐれに論理の飛躍を検出して遊び心を演出
 */

class RandomDetector {
  constructor(rate = 5) {
    this.rate = rate; // 検出確率（パーセント）
    
    // 気まぐれメッセージ集
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
      '気のせいかもしれませんが...',
      'ここは見逃せない予感',
      '理由はわからないけど気になる',
      'センサーが反応しました',
      'なんか、ね？',
      'ここだけは押さずにいられない'
    ];
    
    // 気まぐれスタイル
    this.whimsicalStyles = [
      { emoji: '🌟', color: '#FFD700' },
      { emoji: '✨', color: '#87CEEB' },
      { emoji: '💫', color: '#DDA0DD' },
      { emoji: '🎲', color: '#FF6347' },
      { emoji: '🎪', color: '#FF1493' }
    ];
  }
  
  /**
   * ランダム検出を実行
   */
  detect(text) {
    if (this.rate <= 0) {
      return [];
    }
    
    const sentences = this.extractSentences(text);
    const detections = [];
    
    for (const sentence of sentences) {
      // 確率判定
      if (this.shouldDetect()) {
        const detection = this.createRandomDetection(sentence, text);
        detections.push(detection);
      }
    }
    
    return detections;
  }
  
  /**
   * 検出すべきかを確率で判定
   */
  shouldDetect() {
    return Math.random() * 100 < this.rate;
  }
  
  /**
   * テキストから文を抽出
   */
  extractSentences(text) {
    const sentences = [];
    const parts = text.split(/[。!?]/);
    
    let position = 0;
    for (const part of parts) {
      const trimmed = part.trim();
      if (trimmed.length >= 10) { // 10文字以上の文のみ
        sentences.push({
          text: trimmed,
          position: text.indexOf(part, position),
          length: part.length
        });
        position += part.length + 1;
      }
    }
    
    return sentences;
  }
  
  /**
   * ランダム検出データを生成
   */
  createRandomDetection(sentence, fullText) {
    const message = this.getRandomMessage();
    const style = this.getRandomStyle();
    const score = this.calculateRandomScore();
    
    return {
      text: this.truncateText(sentence.text, 80),
      score: score,
      type: 'random',
      position: sentence.position,
      reason: message,
      style: style,
      isWhimsical: true
    };
  }
  
  /**
   * ランダムメッセージを取得
   */
  getRandomMessage() {
    const index = Math.floor(Math.random() * this.whimsicalMessages.length);
    return this.whimsicalMessages[index];
  }
  
  /**
   * ランダムスタイルを取得
   */
  getRandomStyle() {
    const index = Math.floor(Math.random() * this.whimsicalStyles.length);
    return this.whimsicalStyles[index];
  }
  
  /**
   * ランダムスコアを計算
   */
  calculateRandomScore() {
    // 30-60%の範囲でランダム
    return Math.round(30 + Math.random() * 30);
  }
  
  /**
   * テキストを切り詰め
   */
  truncateText(text, maxLength) {
    if (text.length <= maxLength) {
      return text;
    }
    return text.substring(0, maxLength) + '...';
  }
  
  /**
   * 特定キーワードで検出確率を上げる
   */
  detectWithBias(text, keywords = []) {
    const sentences = this.extractSentences(text);
    const detections = [];
    
    for (const sentence of sentences) {
      let biasedRate = this.rate;
      
      // キーワードが含まれる場合は確率を上げる
      for (const keyword of keywords) {
        if (sentence.text.includes(keyword)) {
          biasedRate *= 2; // 確率2倍
          break;
        }
      }
      
      if (Math.random() * 100 < biasedRate) {
        const detection = this.createRandomDetection(sentence, text);
        detection.reason = `${detection.reason}（キーワード反応）`;
        detections.push(detection);
      }
    }
    
    return detections;
  }
  
  /**
   * 時間帯によって検出確率を変える（おまけ機能）
   */
  detectByTime(text) {
    const hour = new Date().getHours();
    let timeRate = this.rate;
    
    // 深夜（23-5時）は確率アップ（眠くて判断が甘くなる演出）
    if (hour >= 23 || hour <= 5) {
      timeRate *= 1.5;
    }
    
    // 昼休み（12-13時）も確率アップ（リラックスモード）
    if (hour >= 12 && hour <= 13) {
      timeRate *= 1.3;
    }
    
    const tempDetector = new RandomDetector(timeRate);
    return tempDetector.detect(text);
  }
  
  /**
   * コンボシステム（連続検出でメッセージが変わる）
   */
  createComboDetection(comboCount) {
    const comboMessages = [
      '調子に乗ってきました🪽',
      'ノッてます！',
      '止まらない〜✨',
      'コンボ継続中！',
      'やめられない止まらない',
      '気分が高まってきた',
      'もうダメ、楽しすぎる',
      '完全に勢いです'
    ];
    
    if (comboCount >= comboMessages.length) {
      return comboMessages[comboMessages.length - 1];
    }
    
    return comboMessages[comboCount];
  }
}

/**
 * Apps Scriptからランダム検出を実行
 */
function executeRandomDetection(text, rate) {
  const detector = new RandomDetector(rate || 5);
  return detector.detect(text);
}

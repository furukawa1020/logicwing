/**
 * 論理の飛躍検出エンジン (Logic Leap Detector)
 * TypeScript版 for Word Add-in
 */

export interface LogicLeap {
  text: string;
  score: number;
  type: 'pattern' | 'polished' | 'random' | 'ai' | 'manual';
  position: number;
  reason: string;
}

interface LeapPattern {
  pattern: RegExp;
  weight: number;
}

export class LogicLeapDetector {
  private logicalConnectors: string[];
  private leapPatterns: LeapPattern[];
  private tooPolishedPatterns: RegExp[];

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

  public detect(text: string, randomRate: number = 5): LogicLeap[] {
    const leaps: LogicLeap[] = [];

    leaps.push(...this.detectByPatterns(text));
    leaps.push(...this.detectTooPolished(text));

    if (randomRate > 0) {
      const randomDetector = new RandomDetector(randomRate);
      leaps.push(...randomDetector.detect(text));
    }

    return this.deduplicateAndScore(leaps);
  }

  private detectByPatterns(text: string): LogicLeap[] {
    const leaps: LogicLeap[] = [];

    for (const { pattern, weight } of this.leapPatterns) {
      const regex = new RegExp(pattern.source, pattern.flags);
      let match: RegExpExecArray | null;

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

  private detectTooPolished(text: string): LogicLeap[] {
    const leaps: LogicLeap[] = [];

    for (const pattern of this.tooPolishedPatterns) {
      const regex = new RegExp(pattern.source, pattern.flags);
      let match: RegExpExecArray | null;

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

  private calculateContextScore(text: string, position: number, length: number): number {
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

  private countConnectors(text: string): number {
    let count = 0;
    for (const connector of this.logicalConnectors) {
      if (text.includes(connector)) count++;
    }
    return count;
  }

  private getReasonForPattern(text: string): string {
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

  private deduplicateAndScore(leaps: LogicLeap[]): LogicLeap[] {
    const merged: LogicLeap[] = [];
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

/**
 * ランダム検出器 (Random Detector)
 */
class RandomDetector {
  private rate: number;
  private whimsicalMessages: string[];

  constructor(rate: number) {
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

  public detect(text: string): LogicLeap[] {
    if (this.rate <= 0) {
      return [];
    }

    const sentences = this.extractSentences(text);
    const detections: LogicLeap[] = [];

    for (const sentence of sentences) {
      if (Math.random() * 100 < this.rate) {
        const detection = this.createRandomDetection(sentence);
        detections.push(detection);
      }
    }

    return detections;
  }

  private extractSentences(text: string): Array<{ text: string; position: number }> {
    const sentences: Array<{ text: string; position: number }> = [];
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

  private createRandomDetection(sentence: { text: string; position: number }): LogicLeap {
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

  private getRandomMessage(): string {
    const index = Math.floor(Math.random() * this.whimsicalMessages.length);
    return this.whimsicalMessages[index];
  }
}

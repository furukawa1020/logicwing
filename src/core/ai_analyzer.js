/**
 * AI補助検出モジュール (ai_analyzer.js)
 * OpenAI GPT APIを使用した高度な文脈分析
 */

class AIAnalyzer {
  constructor(apiKey) {
    this.apiKey = apiKey;
    this.apiEndpoint = 'https://api.openai.com/v1/chat/completions';
    this.model = 'gpt-3.5-turbo';
  }
  
  /**
   * テキストの論理飛躍をAI分析
   */
  analyzeLogicalLeaps(text) {
    if (!this.apiKey) {
      throw new Error('API Keyが設定されていません');
    }
    
    const prompt = this.buildPrompt(text);
    const response = this.callAPI(prompt);
    
    return this.parseResponse(response);
  }
  
  /**
   * プロンプトを構築
   */
  buildPrompt(text) {
    return `あなたは論理的思考の専門家です。以下の文章を分析し、論理の飛躍がある箇所を特定してください。

【分析対象の文章】
${text}

【出力形式】
JSON形式で以下の構造で返してください:
{
  "leaps": [
    {
      "text": "飛躍している文章の抜粋（最大100文字）",
      "position": 文章内の開始位置（概算でOK）,
      "score": 飛躍度（0-100の整数）,
      "reason": "飛躍の理由（簡潔に）"
    }
  ]
}

【判定基準】
1. 前提が不明確なまま結論に至っている
2. 因果関係が不十分
3. 論理的なステップが省略されている
4. 断定的すぎる表現が使われている
5. 文脈の繋がりが不自然

飛躍が見つからない場合は空の配列を返してください。
最大5箇所まで検出してください。`;
  }
  
  /**
   * OpenAI APIを呼び出し
   */
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
            {
              role: 'system',
              content: 'あなたは論理的思考の専門家です。文章の論理構造を分析し、論理の飛躍を指摘します。'
            },
            {
              role: 'user',
              content: prompt
            }
          ],
          temperature: 0.3,
          max_tokens: 1000
        }),
        muteHttpExceptions: true
      });
      
      const result = JSON.parse(response.getContentText());
      
      if (result.error) {
        throw new Error(`API Error: ${result.error.message}`);
      }
      
      return result.choices[0].message.content;
      
    } catch (e) {
      Logger.log(`API呼び出しエラー: ${e.message}`);
      throw e;
    }
  }
  
  /**
   * APIレスポンスをパース
   */
  parseResponse(responseText) {
    try {
      // レスポンスからJSONを抽出
      const jsonMatch = responseText.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        return [];
      }
      
      const data = JSON.parse(jsonMatch[0]);
      
      if (!data.leaps || !Array.isArray(data.leaps)) {
        return [];
      }
      
      // AIの検出結果を標準形式に変換
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
  
  /**
   * 文章を分割してバッチ分析
   */
  analyzeBatch(text, chunkSize = 2000) {
    const chunks = this.splitText(text, chunkSize);
    const allLeaps = [];
    
    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const leaps = this.analyzeLogicalLeaps(chunk.text);
      
      // 位置オフセットを調整
      leaps.forEach(leap => {
        leap.position += chunk.offset;
        allLeaps.push(leap);
      });
      
      // API制限を考慮して待機
      if (i < chunks.length - 1) {
        Utilities.sleep(1000);
      }
    }
    
    return allLeaps;
  }
  
  /**
   * テキストを適切なサイズに分割
   */
  splitText(text, chunkSize) {
    const chunks = [];
    const sentences = text.split(/([。!?])/);
    
    let currentChunk = '';
    let offset = 0;
    
    for (let i = 0; i < sentences.length; i += 2) {
      const sentence = sentences[i] + (sentences[i + 1] || '');
      
      if (currentChunk.length + sentence.length > chunkSize && currentChunk.length > 0) {
        chunks.push({
          text: currentChunk,
          offset: offset
        });
        offset += currentChunk.length;
        currentChunk = sentence;
      } else {
        currentChunk += sentence;
      }
    }
    
    if (currentChunk.length > 0) {
      chunks.push({
        text: currentChunk,
        offset: offset
      });
    }
    
    return chunks;
  }
  
  /**
   * API接続をテスト
   */
  testConnection() {
    try {
      const response = this.analyzeLogicalLeaps('これはテストです。問題ありません。');
      return {
        success: true,
        message: 'API接続成功'
      };
    } catch (e) {
      return {
        success: false,
        message: e.message
      };
    }
  }
}

/**
 * Apps ScriptからAI分析を実行
 */
function analyzeWithAI(text) {
  const settings = getSettings();
  
  if (!settings.enableAI || !settings.aiApiKey) {
    return {
      success: false,
      message: 'AI検出が無効、またはAPIキーが未設定です'
    };
  }
  
  try {
    const analyzer = new AIAnalyzer(settings.aiApiKey);
    const leaps = analyzer.analyzeLogicalLeaps(text);
    
    return {
      success: true,
      leaps: leaps
    };
  } catch (e) {
    return {
      success: false,
      message: e.message
    };
  }
}

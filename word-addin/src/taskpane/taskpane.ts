/*
 * 論理の飛躍スタンプ拡張機能（ロンリちゃん🪽）
 * メイン処理 - taskpane.ts
 */

import { LogicLeapDetector, LogicLeap } from '../detector/LogicLeapDetector';
import { StorageManager, Settings } from '../detector/StorageManager';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    document.getElementById('detect-btn').onclick = detectLogicalLeaps;
    document.getElementById('manual-stamp-btn').onclick = manualStamp;
    document.getElementById('show-history-btn').onclick = showHistory;
    document.getElementById('settings-btn').onclick = showSettings;
    
    loadSettings();
  }
});

let currentSettings: Settings;

/**
 * 設定を読み込み
 */
async function loadSettings() {
  currentSettings = await StorageManager.getSettings();
  console.log('設定読み込み:', currentSettings);
}

/**
 * 論理の飛躍を検出
 */
async function detectLogicalLeaps() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load('text');
      await context.sync();

      const text = body.text;

      if (!text || text.trim().length === 0) {
        showMessage('文章が入力されていません', 'エラー');
        return;
      }

      const detector = new LogicLeapDetector();
      const leaps = detector.detect(text, currentSettings.randomRate);

      if (leaps.length === 0) {
        showMessage('論理の飛躍は検出されませんでした！✨<br>文章は論理的に整っています。', '優秀です');
        return;
      }

      // アニメーション表示
      showStampingAnimation(leaps.length);
      
      // 文書内に印面を押す（順番に）
      await stampOnDocument(context, leaps);
      
      // 履歴に保存
      for (const leap of leaps) {
        await StorageManager.saveStampData(leap);
      }
      
      // 完了表示
      setTimeout(() => {
        showMessage(`🪽 ${leaps.length}箇所に押印しました！`, '完了');
      }, 500);
    });
  } catch (error) {
    console.error('検出エラー:', error);
    showMessage('エラーが発生しました: ' + error.message, 'エラー');
  }
}

/**
 * 手動で押印
 */
async function manualStamp() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load('text');
      await context.sync();

      const text = selection.text;

      if (!text || text.trim().length === 0) {
        showMessage('テキストを選択してください', 'エラー');
        return;
      }

      const leap: LogicLeap = {
        text: text.substring(0, 100),
        score: 100,
        type: 'manual',
        position: 0,
        reason: '手動押印'
      };

      // アニメーション表示
      showStampingAnimation(1);

      // 少し待ってから押印
      await new Promise(resolve => setTimeout(resolve, 400));

      // 選択範囲の後ろに印面を挿入
      const stampImage = await loadStampImage();
      const inlinePicture = selection.insertInlinePictureFromBase64(stampImage, Word.InsertLocation.after);
      inlinePicture.height = 60;
      inlinePicture.width = 60;
      inlinePicture.lockAspectRatio = true;
      
      await context.sync();
      await StorageManager.saveStampData(leap);
      
      if (currentSettings.enableSound) {
        playStampSound();
      }
      
      // 完了表示
      setTimeout(() => {
        showMessage('手動押印しました！', '完了');
      }, 300);
    });
  } catch (error) {
    console.error('手動押印エラー:', error);
    showMessage('エラーが発生しました: ' + error.message, 'エラー');
  }
}

/**
 * 検出結果を表示（Task Pane用・シンプル版）
 */
function displayLeaps(leaps: LogicLeap[]) {
  const resultsDiv = document.getElementById('results');
  resultsDiv.innerHTML = '';
  resultsDiv.style.display = 'block';

  const title = document.createElement('h3');
  title.textContent = `🪽 ${leaps.length}箇所に押印しました`;
  title.style.color = '#333';
  title.style.fontSize = '16px';
  title.style.marginBottom = '15px';
  title.style.fontWeight = 'normal';
  resultsDiv.appendChild(title);

  // シンプルなリスト表示のみ
  const listDiv = document.createElement('div');
  listDiv.style.display = 'flex';
  listDiv.style.flexDirection = 'column';
  listDiv.style.gap = '10px';

  leaps.forEach((leap, index) => {
    const item = document.createElement('div');
    item.style.padding = '10px';
    item.style.background = '#f8f9fa';
    item.style.borderRadius = '4px';
    item.style.borderLeft = '3px solid #ff4081';
    item.style.fontSize = '13px';

    const header = document.createElement('div');
    header.style.display = 'flex';
    header.style.justifyContent = 'space-between';
    header.style.marginBottom = '5px';
    header.style.fontSize = '12px';
    header.style.color = '#666';

    const num = document.createElement('span');
    num.textContent = `#${index + 1}`;
    num.style.fontWeight = 'bold';

    const score = document.createElement('span');
    score.textContent = `${leap.score}点`;
    score.style.color = leap.score >= 70 ? '#d32f2f' : leap.score >= 50 ? '#f57c00' : '#1976d2';
    score.style.fontWeight = 'bold';

    header.appendChild(num);
    header.appendChild(score);

    const text = document.createElement('div');
    text.textContent = leap.text.length > 50 ? leap.text.substring(0, 50) + '...' : leap.text;
    text.style.color = '#333';
    text.style.marginBottom = '3px';

    const reason = document.createElement('div');
    reason.textContent = leap.reason;
    reason.style.fontSize = '11px';
    reason.style.color = '#888';

    item.appendChild(header);
    item.appendChild(text);
    item.appendChild(reason);

    listDiv.appendChild(item);
  });

  resultsDiv.appendChild(listDiv);
}

/**
 * 押印アニメーションを表示
 */
function showStampingAnimation(count: number) {
  const resultsDiv = document.getElementById('results');
  resultsDiv.innerHTML = '';
  resultsDiv.style.display = 'block';

  const animationContainer = document.createElement('div');
  animationContainer.style.cssText = `
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 40px 20px;
    min-height: 200px;
  `;

  const stampImg = document.createElement('img');
  stampImg.src = '../../assets/論理の翼.png';
  stampImg.style.cssText = `
    width: 120px;
    height: 120px;
    animation: stampFlying 0.8s ease-in-out infinite;
  `;

  const message = document.createElement('div');
  message.style.cssText = `
    margin-top: 20px;
    font-size: 16px;
    font-weight: 600;
    color: #333;
    text-align: center;
  `;
  message.textContent = `押印中... (${count}箇所)`;

  animationContainer.appendChild(stampImg);
  animationContainer.appendChild(message);
  resultsDiv.appendChild(animationContainer);

  // アニメーションCSSを追加
  if (!document.getElementById('stamp-animation-style')) {
    const style = document.createElement('style');
    style.id = 'stamp-animation-style';
    style.textContent = `
      @keyframes stampFlying {
        0% { transform: translateY(0) rotate(0deg); }
        50% { transform: translateY(-15px) rotate(-5deg); }
        100% { transform: translateY(0) rotate(0deg); }
      }
    `;
    document.head.appendChild(style);
  }
}

/**
 * 文書内にスタンプを押す（メイン処理）
 */
async function stampOnDocument(context: Word.RequestContext, leaps: LogicLeap[]) {
  const body = context.document.body;
  const stampImage = await loadStampImage();
  
  let stampedCount = 0;

  for (const leap of leaps) {
    try {
      // 該当テキストを検索
      const searchResults = body.search(leap.text, {
        matchCase: false,
        matchWholeWord: false,
        matchWildcards: false
      });
      
      searchResults.load('items');
      await context.sync();

      if (searchResults.items.length > 0) {
        // 最初の検出箇所に印面を挿入
        const range = searchResults.items[0];
        
        // テキストをピンク色でハイライト
        range.font.highlightColor = '#FFE4E9';
        range.font.color = '#ff4081';
        
        // 範囲の後ろに印面画像を挿入
        const inlinePicture = range.insertInlinePictureFromBase64(
          stampImage, 
          Word.InsertLocation.after
        );
        
        // サイズを適切に設定（実際の印鑑サイズに近く）
        inlinePicture.height = 60;
        inlinePicture.width = 60;
        inlinePicture.lockAspectRatio = true;
        
        // スペースを追加
        range.insertText(' ', Word.InsertLocation.after);
        
        stampedCount++;
        await context.sync();
        
        // 押印効果音
        if (currentSettings.enableSound) {
          playStampSound();
        }
        
        // 次の押印まで少し待つ（視覚効果）
        await new Promise(resolve => setTimeout(resolve, 300));
      }
    } catch (error) {
      console.warn(`スタンプ押印失敗 for "${leap.text}":`, error);
    }
  }
}

/**
 * 印面画像をBase64で読み込み
 */
async function loadStampImage(): Promise<string> {
  // 印面.pngをBase64に変換
  const response = await fetch('../../assets/印面.png');
  const blob = await response.blob();
  
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const base64 = reader.result as string;
      // data:image/png;base64, の部分を削除
      const base64Data = base64.split(',')[1];
      resolve(base64Data);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

/**
 * スコアに応じた色を取得
 */
function getScoreColor(score: number): string {
  if (score >= 80) return '#ff4444';
  if (score >= 60) return '#ff9944';
  if (score >= 40) return '#ffcc44';
  return '#44ccff';
}

/**
 * HTMLエスケープ
 */
function escapeHtml(text: string): string {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

/**
 * サウンド再生（押印音）
 */
function playStampSound() {
  try {
    const audioContext = new (window.AudioContext || (window as any).webkitAudioContext)();
    
    // 押印の「ポンッ」という音を再現
    // 低音と高音を組み合わせる
    const oscillator1 = audioContext.createOscillator();
    const oscillator2 = audioContext.createOscillator();
    const gainNode = audioContext.createGain();

    oscillator1.connect(gainNode);
    oscillator2.connect(gainNode);
    gainNode.connect(audioContext.destination);

    // 低音（ドンという音）
    oscillator1.frequency.value = 150;
    oscillator1.type = 'triangle';

    // 高音（パンという音）
    oscillator2.frequency.value = 600;
    oscillator2.type = 'sine';

    // 音量を急激に下げる（短く鋭い音）
    gainNode.gain.setValueAtTime(0.4, audioContext.currentTime);
    gainNode.gain.exponentialRampToValueAtTime(0.01, audioContext.currentTime + 0.08);

    oscillator1.start(audioContext.currentTime);
    oscillator2.start(audioContext.currentTime);
    oscillator1.stop(audioContext.currentTime + 0.08);
    oscillator2.stop(audioContext.currentTime + 0.08);
  } catch (error) {
    console.warn('音声再生エラー:', error);
  }
}

/**
 * メッセージ表示
 */
function showMessage(message: string, title: string = 'お知らせ') {
  const resultsDiv = document.getElementById('results');
  resultsDiv.innerHTML = `
    <div class="message-box">
      <h3>${escapeHtml(title)}</h3>
      <p>${message}</p>
    </div>
  `;
  resultsDiv.style.display = 'block';
}

/**
 * 履歴を表示
 */
async function showHistory() {
  const history = await StorageManager.getStampHistory();
  const stats = await StorageManager.getStatistics();

  const resultsDiv = document.getElementById('results');
  resultsDiv.innerHTML = '';
  resultsDiv.style.display = 'block';

  const title = document.createElement('h3');
  title.textContent = '🪽 押印履歴';
  title.style.color = '#ff4081';
  resultsDiv.appendChild(title);

  // 統計情報
  const statsDiv = document.createElement('div');
  statsDiv.className = 'stats-box';
  statsDiv.innerHTML = `
    <h4>統計情報</h4>
    <p>総押印数: ${stats.totalStamps}回</p>
    <p>平均スコア: ${stats.averageScore}点</p>
  `;
  resultsDiv.appendChild(statsDiv);

  // 履歴リスト
  if (history.length === 0) {
    resultsDiv.innerHTML += '<p>履歴はありません</p>';
    return;
  }

  const listDiv = document.createElement('div');
  listDiv.className = 'history-list';

  history.slice(0, 20).forEach((entry, index) => {
    const item = document.createElement('div');
    item.className = 'history-item';

    const date = new Date(entry.timestamp);
    const dateStr = `${date.getMonth() + 1}/${date.getDate()} ${date.getHours()}:${String(date.getMinutes()).padStart(2, '0')}`;

    item.innerHTML = `
      <div class="history-header">
        <span class="history-date">${dateStr}</span>
        <span class="history-score">${entry.score}点</span>
      </div>
      <div class="history-text">${escapeHtml(entry.text.substring(0, 50))}...</div>
      <div class="history-reason">${escapeHtml(entry.reason)}</div>
    `;

    listDiv.appendChild(item);
  });

  resultsDiv.appendChild(listDiv);
}

/**
 * 設定を表示
 */
async function showSettings() {
  const settings = await StorageManager.getSettings();

  const resultsDiv = document.getElementById('results');
  resultsDiv.innerHTML = '';
  resultsDiv.style.display = 'block';

  const title = document.createElement('h3');
  title.textContent = '⚙️ 設定';
  title.style.color = '#ff4081';
  resultsDiv.appendChild(title);

  const formDiv = document.createElement('div');
  formDiv.className = 'settings-form';
  formDiv.innerHTML = `
    <div class="setting-item">
      <label>ランダム検出率 (%)</label>
      <input type="range" id="random-rate" min="0" max="20" value="${settings.randomRate}" />
      <span id="random-rate-value">${settings.randomRate}%</span>
    </div>
    
    <div class="setting-item">
      <label>
        <input type="checkbox" id="enable-sound" ${settings.enableSound ? 'checked' : ''} />
        効果音を有効にする
      </label>
    </div>
    
    <div class="setting-item">
      <button id="save-settings-btn" class="primary-btn">保存</button>
    </div>
  `;

  resultsDiv.appendChild(formDiv);

  // イベントリスナー
  const rangeInput = document.getElementById('random-rate') as HTMLInputElement;
  const rangeValue = document.getElementById('random-rate-value');
  rangeInput.oninput = () => {
    rangeValue.textContent = `${rangeInput.value}%`;
  };

  document.getElementById('save-settings-btn').onclick = async () => {
    const newSettings: Settings = {
      ...settings,
      randomRate: parseInt(rangeInput.value),
      enableSound: (document.getElementById('enable-sound') as HTMLInputElement).checked
    };

    await StorageManager.saveSettings(newSettings);
    currentSettings = newSettings;
    showMessage('設定を保存しました', '完了');
  };
}

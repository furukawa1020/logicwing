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

      // 検出結果を表示
      displayLeaps(leaps);
      
      // 履歴に保存
      for (const leap of leaps) {
        await StorageManager.saveStampData(leap);
      }
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

      displayLeaps([leap]);
      await StorageManager.saveStampData(leap);
    });
  } catch (error) {
    console.error('手動押印エラー:', error);
    showMessage('エラーが発生しました: ' + error.message, 'エラー');
  }
}

/**
 * 検出結果を表示
 */
function displayLeaps(leaps: LogicLeap[]) {
  const resultsDiv = document.getElementById('results');
  resultsDiv.innerHTML = '';
  resultsDiv.style.display = 'block';

  const title = document.createElement('h3');
  title.textContent = `🪽 ${leaps.length}箇所検出しました`;
  title.style.color = '#ff4081';
  title.style.marginBottom = '20px';
  resultsDiv.appendChild(title);

  // アニメーションコンテナ
  const animationDiv = document.createElement('div');
  animationDiv.className = 'stamp-animation-container';
  animationDiv.innerHTML = `
    <div class="ronri-chan">
      <img src="../../assets/印面.png" class="stamp-body" alt="ロンリちゃん" />
      <img src="../../assets/論理の翼.png" class="wing wing-left" alt="左翼" />
      <img src="../../assets/論理の翼.png" class="wing wing-right" alt="右翼" />
    </div>
  `;
  resultsDiv.appendChild(animationDiv);

  // 検出リスト
  const listDiv = document.createElement('div');
  listDiv.className = 'leap-list';

  leaps.forEach((leap, index) => {
    const item = document.createElement('div');
    item.className = 'leap-item';
    item.style.animationDelay = `${index * 0.1}s`;

    const scoreColor = getScoreColor(leap.score);

    item.innerHTML = `
      <div class="leap-header">
        <span class="leap-number">#${index + 1}</span>
        <span class="leap-score" style="background: ${scoreColor}">${leap.score}点</span>
      </div>
      <div class="leap-text">${escapeHtml(leap.text)}</div>
      <div class="leap-reason">💡 ${escapeHtml(leap.reason)}</div>
    `;

    listDiv.appendChild(item);
  });

  resultsDiv.appendChild(listDiv);

  // サウンド再生
  if (currentSettings.enableSound) {
    playStampSound();
  }
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
 * サウンド再生
 */
function playStampSound() {
  // 簡易的なビープ音
  const audioContext = new (window.AudioContext || (window as any).webkitAudioContext)();
  const oscillator = audioContext.createOscillator();
  const gainNode = audioContext.createGain();

  oscillator.connect(gainNode);
  gainNode.connect(audioContext.destination);

  oscillator.frequency.value = 800;
  oscillator.type = 'sine';

  gainNode.gain.setValueAtTime(0.3, audioContext.currentTime);
  gainNode.gain.exponentialRampToValueAtTime(0.01, audioContext.currentTime + 0.1);

  oscillator.start(audioContext.currentTime);
  oscillator.stop(audioContext.currentTime + 0.1);
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

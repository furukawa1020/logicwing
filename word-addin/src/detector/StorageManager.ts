/**
 * データ管理 (Storage Manager)
 * Word Add-in用のローカルストレージ管理
 */

import { LogicLeap } from './LogicLeapDetector';

export interface StampEntry {
  text: string;
  score: number;
  type: string;
  reason: string;
  timestamp: string;
}

export interface Settings {
  theme: 'logical_leap' | 'cute' | 'serious';
  randomRate: number;
  enableSound: boolean;
  enableAI: boolean;
  aiApiKey: string;
}

export class StorageManager {
  private static readonly HISTORY_KEY = 'ronri_history';
  private static readonly SETTINGS_KEY = 'ronri_settings';
  private static readonly MAX_HISTORY = 100;

  /**
   * スタンプデータを保存
   */
  public static async saveStampData(leapData: LogicLeap): Promise<boolean> {
    try {
      let history = await this.getStampHistory();

      const stampEntry: StampEntry = {
        text: leapData.text,
        score: leapData.score,
        type: leapData.type,
        reason: leapData.reason,
        timestamp: new Date().toISOString()
      };

      history.unshift(stampEntry);

      if (history.length > this.MAX_HISTORY) {
        history = history.slice(0, this.MAX_HISTORY);
      }

      localStorage.setItem(this.HISTORY_KEY, JSON.stringify(history));
      return true;
    } catch (e) {
      console.error('スタンプデータの保存エラー:', e);
      return false;
    }
  }

  /**
   * スタンプ履歴を取得
   */
  public static async getStampHistory(): Promise<StampEntry[]> {
    try {
      const stored = localStorage.getItem(this.HISTORY_KEY);

      if (!stored) {
        return [];
      }

      return JSON.parse(stored);
    } catch (e) {
      console.error('履歴の取得エラー:', e);
      return [];
    }
  }

  /**
   * 設定を保存
   */
  public static async saveSettings(settings: Settings): Promise<boolean> {
    try {
      localStorage.setItem(this.SETTINGS_KEY, JSON.stringify(settings));
      return true;
    } catch (e) {
      console.error('設定の保存エラー:', e);
      return false;
    }
  }

  /**
   * 設定を取得
   */
  public static async getSettings(): Promise<Settings> {
    try {
      const stored = localStorage.getItem(this.SETTINGS_KEY);

      const defaultSettings: Settings = {
        theme: 'logical_leap',
        randomRate: 5,
        enableSound: true,
        enableAI: false,
        aiApiKey: ''
      };

      if (!stored) {
        return defaultSettings;
      }

      return { ...defaultSettings, ...JSON.parse(stored) };
    } catch (e) {
      console.error('設定の取得エラー:', e);
      return {
        theme: 'logical_leap',
        randomRate: 5,
        enableSound: true,
        enableAI: false,
        aiApiKey: ''
      };
    }
  }

  /**
   * 履歴をクリア
   */
  public static async clearHistory(): Promise<boolean> {
    try {
      localStorage.removeItem(this.HISTORY_KEY);
      return true;
    } catch (e) {
      console.error('履歴のクリアエラー:', e);
      return false;
    }
  }

  /**
   * 統計情報を取得
   */
  public static async getStatistics(): Promise<{
    totalStamps: number;
    byType: Record<string, number>;
    averageScore: number;
  }> {
    const history = await this.getStampHistory();

    const byType: Record<string, number> = {};
    let totalScore = 0;

    for (const entry of history) {
      byType[entry.type] = (byType[entry.type] || 0) + 1;
      totalScore += entry.score;
    }

    return {
      totalStamps: history.length,
      byType,
      averageScore: history.length > 0 ? Math.round(totalScore / history.length) : 0
    };
  }
}

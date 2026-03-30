import { useState, useEffect } from 'react';
import type { AppSettings, GeminiModel } from '../types';
import { initGemini } from '../services/gemini';

const STORAGE_KEY = 'ppt-ai-settings';

const DEFAULT_SETTINGS: AppSettings = {
  apiKey: '',
  model: 'gemini-2.0-flash',
  autoContext: true,
  customPrompts: [],
};


export function useSettings() {
  const [settings, setSettings] = useState<AppSettings>(() => {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) {
        return { ...DEFAULT_SETTINGS, ...JSON.parse(stored) };
      }
    } catch {
      // ignore
    }
    return DEFAULT_SETTINGS;
  });

  // 설정 저장 및 Gemini 초기화
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
    if (settings.apiKey) {
      initGemini(settings.apiKey, settings.model);
    }
  }, [settings]);

  const updateSettings = (updates: Partial<AppSettings>) => {
    setSettings((prev) => ({ ...prev, ...updates }));
  };

  const saveApiKey = (apiKey: string) => {
    updateSettings({ apiKey });
  };

  const setModel = (model: GeminiModel) => {
    updateSettings({ model });
  };

  const hasApiKey = settings.apiKey.length > 0;

  return {
    settings,
    updateSettings,
    saveApiKey,
    setModel,
    hasApiKey,
  };
}

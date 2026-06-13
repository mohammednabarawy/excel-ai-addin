export interface AppSettings {
  provider: 'gemini' | 'lmstudio' | 'opencode';
  geminiApiKey: string;
  geminiModel?: string;
  lmStudioUrl: string;
  lmStudioModel?: string;
  opencodeApiKey: string;
  opencodeModel?: string;
  opencodeUseProxy?: boolean;
}

const DEFAULT_SETTINGS: AppSettings = {
  provider: 'lmstudio',
  geminiApiKey: '',
  geminiModel: 'gemini-1.5-pro-latest',
  lmStudioUrl: 'http://localhost:1234/v1',
  lmStudioModel: 'local-model',
  opencodeApiKey: '',
  opencodeModel: 'minimax-m2.5-free',
  opencodeUseProxy: false,
};

export const loadSettings = (): AppSettings => {
  try {
    const data = localStorage.getItem('antigravity_settings');
    if (data) {
      return { ...DEFAULT_SETTINGS, ...JSON.parse(data) };
    }
  } catch (e) {
    console.error('Failed to load settings', e);
  }
  return DEFAULT_SETTINGS;
};

export const saveSettings = (settings: AppSettings) => {
  try {
    localStorage.setItem('antigravity_settings', JSON.stringify(settings));
  } catch (e) {
    console.error('Failed to save settings', e);
  }
};

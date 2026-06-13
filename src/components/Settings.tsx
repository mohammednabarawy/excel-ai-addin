import { useState } from 'react';
import { 
  Button, 
  Input, 
  Select, 
  Field, 
  makeStyles,
  tokens
} from '@fluentui/react-components';
import { loadSettings, saveSettings, type AppSettings } from '../utils/storage';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    padding: '16px',
  },
  header: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
  }
});

interface SettingsProps {
  onBack: () => void;
}

export const Settings: React.FC<SettingsProps> = ({ onBack }) => {
  const styles = useStyles();
  const [settings, setSettings] = useState<AppSettings>(loadSettings());
  const [availableModels, setAvailableModels] = useState<string[]>([]);
  const [isFetchingModels, setIsFetchingModels] = useState(false);

  const fetchGeminiModels = async () => {
    if (!settings.geminiApiKey) return;
    setIsFetchingModels(true);
    try {
      const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${settings.geminiApiKey}`);
      const data = await response.json();
      if (data.models) {
        const models = data.models
          .filter((m: any) => m.supportedGenerationMethods?.includes('generateContent'))
          .map((m: any) => m.name.replace('models/', ''));
        setAvailableModels(models);
        if (models.length > 0 && !settings.geminiModel) {
          setSettings({ ...settings, geminiModel: models[0] });
        }
      }
    } catch (e) {
      console.error('Failed to fetch models', e);
    } finally {
      setIsFetchingModels(false);
    }
  };

  const handleSave = () => {
    saveSettings(settings);
    onBack();
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>Settings</div>
      
      <Field label="AI Provider">
        <Select 
          value={settings.provider} 
          onChange={(_, data) => setSettings({ ...settings, provider: data.value as any })}
        >
          <option value="lmstudio">Local LM Studio</option>
          <option value="gemini">Google Gemini API</option>
          <option value="opencode">OpenCode Zen</option>
        </Select>
      </Field>

      {settings.provider === 'lmstudio' && (
        <>
          <Field label="LM Studio URL">
            <Input 
              value={settings.lmStudioUrl} 
              onChange={(_, data) => setSettings({ ...settings, lmStudioUrl: data.value })}
            />
          </Field>
          <Field label="LM Studio Model Name">
            <Input 
              value={settings.lmStudioModel || 'local-model'} 
              onChange={(_, data) => setSettings({ ...settings, lmStudioModel: data.value })}
            />
          </Field>
        </>
      )}

      {settings.provider === 'gemini' && (
        <>
          <Field label="Google Gemini API Key">
            <Input 
              type="password"
              value={settings.geminiApiKey} 
              onChange={(_, data) => setSettings({ ...settings, geminiApiKey: data.value })}
            />
          </Field>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'flex-end' }}>
            <Field label="Gemini Model" style={{ flex: 1 }}>
              <Select 
                value={settings.geminiModel || 'gemini-1.5-pro-latest'} 
                onChange={(_, data) => setSettings({ ...settings, geminiModel: data.value })}
              >
                {availableModels.length > 0 ? (
                  availableModels.map(m => <option key={m} value={m}>{m}</option>)
                ) : (
                  <option value={settings.geminiModel || 'gemini-1.5-pro-latest'}>{settings.geminiModel || 'gemini-1.5-pro-latest'}</option>
                )}
              </Select>
            </Field>
            <Button onClick={fetchGeminiModels} disabled={!settings.geminiApiKey || isFetchingModels}>
              {isFetchingModels ? 'Loading...' : 'Load Models'}
            </Button>
          </div>
        </>
      )}

      {settings.provider === 'opencode' && (
        <Field label="OpenCode Zen Key (optional)">
          <Input 
            type="password"
            value={settings.opencodeApiKey} 
            onChange={(_, data) => setSettings({ ...settings, opencodeApiKey: data.value })}
          />
        </Field>
      )}

      <Button appearance="primary" onClick={handleSave}>Save & Close</Button>
      <Button appearance="subtle" onClick={onBack}>Cancel</Button>
    </div>
  );
};

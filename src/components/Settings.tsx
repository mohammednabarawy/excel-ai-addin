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
  const [availableOpenCodeModels, setAvailableOpenCodeModels] = useState<{id: string, name: string}[]>([]);
  const [isFetchingOpenCodeModels, setIsFetchingOpenCodeModels] = useState(false);
  const [openCodeError, setOpenCodeError] = useState('');

  const fetchOpenCodeModels = async () => {
    if (!settings.opencodeApiKey) return;
    setIsFetchingOpenCodeModels(true);
    setOpenCodeError('');
    try {
      const response = await fetch(`https://opencode.ai/zen/v1/models`, {
        headers: { 'Authorization': `Bearer ${settings.opencodeApiKey}` }
      });
      const data = await response.json();
      if (data.data) {
        const models = data.data.map((m: any) => {
          const isFree = m.id.toLowerCase().includes('free');
          return {
            id: m.id,
            name: isFree ? `${m.id} (Free)` : m.id
          };
        });
        setAvailableOpenCodeModels(models);
        if (models.length > 0 && (!settings.opencodeModel || settings.opencodeModel === 'minimax-m2.5-free')) {
          setSettings({ ...settings, opencodeModel: models[0].id });
        }
      } else {
         setOpenCodeError('Invalid response format from OpenCode');
      }
    } catch (e: any) {
      console.error('Failed to fetch OpenCode models', e);
      if (e.message?.includes('Failed to fetch')) {
        setOpenCodeError('CORS Error: OpenCode blocked the request. This usually means your API Key is invalid or missing billing details.');
      } else {
        setOpenCodeError(e.message || 'Unknown error fetching models');
      }
    } finally {
      setIsFetchingOpenCodeModels(false);
    }
  };

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
        <>
          <Field label="OpenCode Zen Key (optional)">
            <Input 
              type="password"
              value={settings.opencodeApiKey} 
              onChange={(_, data) => setSettings({ ...settings, opencodeApiKey: data.value })}
            />
          </Field>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'flex-end' }}>
            <Field label="OpenCode Model" style={{ flex: 1 }}>
              <Select 
                value={settings.opencodeModel || 'minimax-m2.5-free'} 
                onChange={(_, data) => setSettings({ ...settings, opencodeModel: data.value })}
              >
                {availableOpenCodeModels.length > 0 ? (
                  availableOpenCodeModels.map(m => <option key={m.id} value={m.id}>{m.name}</option>)
                ) : (
                  <option value={settings.opencodeModel || 'minimax-m2.5-free'}>{settings.opencodeModel || 'minimax-m2.5-free'}</option>
                )}
              </Select>
            </Field>
            <Button onClick={fetchOpenCodeModels} disabled={!settings.opencodeApiKey || isFetchingOpenCodeModels}>
              {isFetchingOpenCodeModels ? 'Loading...' : 'Load Models'}
            </Button>
          </div>
          {openCodeError && (
            <div style={{ color: 'red', fontSize: '12px', marginTop: '4px' }}>
              {openCodeError}
            </div>
          )}
        </>
      )}

      <Button appearance="primary" onClick={handleSave}>Save & Close</Button>
      <Button appearance="subtle" onClick={onBack}>Cancel</Button>
    </div>
  );
};

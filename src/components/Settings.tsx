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
        <Field label="Google Gemini API Key">
          <Input 
            type="password"
            value={settings.geminiApiKey} 
            onChange={(_, data) => setSettings({ ...settings, geminiApiKey: data.value })}
          />
        </Field>
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

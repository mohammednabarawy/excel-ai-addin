import { useState } from 'react';
import { Button, makeStyles, tokens } from '@fluentui/react-components';
import { Settings } from './components/Settings';
import { Chat } from './components/Chat';

const useStyles = makeStyles({
  appContainer: {
    height: '100vh',
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
  },
  header: {
    padding: '8px 16px',
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  content: {
    flex: 1,
    overflowY: 'auto',
  }
});

function App() {
  const styles = useStyles();
  const [view, setView] = useState<'chat' | 'settings'>('chat');

  return (
    <div className={styles.appContainer}>
      <div className={styles.header}>
        <span style={{ fontWeight: 600 }}>Excel AI</span>
        {view === 'chat' && (
          <Button 
            appearance="transparent" 
            style={{ color: 'white' }}
            onClick={() => setView('settings')}
          >
            Settings
          </Button>
        )}
      </div>

      <div className={styles.content}>
        {view === 'settings' ? (
          <Settings onBack={() => setView('chat')} />
        ) : (
          <Chat />
        )}
      </div>
    </div>
  );
}

export default App;

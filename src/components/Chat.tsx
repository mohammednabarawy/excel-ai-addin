import { useState } from 'react';
import { Button, Input, makeStyles, tokens } from '@fluentui/react-components';
import { loadSettings } from '../utils/storage';
import { getExcelContext, getSpecificRangeData } from '../utils/excelContext';
import { sendToAI, type AIResponse } from '../utils/aiConnector';
import { executeChange, undoLastChange } from '../utils/excelExecution';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    padding: '16px',
    boxSizing: 'border-box'
  },
  chatHistory: {
    flex: 1,
    overflowY: 'auto',
    marginBottom: '16px',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px'
  },
  message: {
    padding: '8px 12px',
    borderRadius: '8px',
    maxWidth: '85%',
  },
  userMessage: {
    alignSelf: 'flex-end',
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand
  },
  aiMessage: {
    alignSelf: 'flex-start',
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground1
  },
  inputArea: {
    display: 'flex',
    gap: '8px'
  },
  approvalBox: {
    marginTop: '8px',
    padding: '8px',
    backgroundColor: tokens.colorPaletteRedBackground1,
    border: `1px solid ${tokens.colorPaletteRedBorder1}`,
    borderRadius: '4px'
  }
});

interface ChatMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  isApprovalPending?: boolean;
  proposedChanges?: any;
}

export const Chat: React.FC = () => {
  const styles = useStyles();
  const [input, setInput] = useState('');
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [canUndo, setCanUndo] = useState(false);

  const handleSend = async () => {
    if (!input.trim() || isProcessing) return;
    
    const userMsg: ChatMessage = { id: Date.now().toString(), role: 'user', content: input };
    setMessages(prev => [...prev, userMsg]);
    setInput('');
    setIsProcessing(true);

    try {
      const settings = loadSettings();
      const excelContext = await getExcelContext();
      
      let historyForAI = messages.map(m => ({ role: m.role, content: m.content }));
      let response: AIResponse = await sendToAI(input, settings, excelContext, historyForAI);

      // Handle AI requesting more data
      let fetchCount = 0;
      while (response.actionType === 'read_more_data' && response.proposedChanges?.rangeAddress && fetchCount < 3) {
        fetchCount++;
        const address = response.proposedChanges.rangeAddress;
        
        // Add visual indicator for the user
        setMessages(prev => [...prev, { id: Date.now().toString() + fetchCount, role: 'assistant', content: `(AI is reading data from ${address}...)` }]);
        
        const extraData = await getSpecificRangeData(address);
        
        historyForAI.push({ role: 'assistant', content: response.messageToUser || `I need to read data from ${address}` });
        historyForAI.push({ role: 'user', content: `Data for ${address}: ${JSON.stringify(extraData)}. Please continue answering the original request.` });
        
        response = await sendToAI("Continue.", settings, excelContext, historyForAI);
      }

      const aiMsg: ChatMessage = { 
        id: (Date.now() + 10).toString(), 
        role: 'assistant', 
        content: response.messageToUser,
        isApprovalPending: response.actionType === 'destructive_edit',
        proposedChanges: response.proposedChanges
      };

      if (response.actionType === 'edit' && response.proposedChanges) {
         await executeChange(response.proposedChanges);
         setCanUndo(true);
      }

      setMessages(prev => [...prev, aiMsg]);
    } catch (e: any) {
       console.error(e);
       let errorMessage = e.message;
       if (errorMessage === 'Failed to fetch') {
         errorMessage = 'Failed to fetch. This usually means the API server is offline, the URL is wrong, or the browser blocked the request due to CORS.';
       }
       setMessages(prev => [...prev, { id: Date.now().toString(), role: 'assistant', content: `Error: ${errorMessage}` }]);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleApprove = async (msgId: string, changes: any) => {
    try {
      if (changes) {
        await executeChange(changes);
        setCanUndo(true);
      }
      setMessages(prev => prev.map(m => m.id === msgId ? { ...m, isApprovalPending: false, content: m.content + ' (Approved & Applied)' } : m));
    } catch (e: any) {
      alert(`Execution failed: ${e.message}`);
    }
  };

  const handleUndo = async () => {
    try {
      await undoLastChange();
      setCanUndo(false);
      setMessages(prev => [...prev, { id: Date.now().toString(), role: 'assistant', content: 'Last action undone.' }]);
    } catch (e: any) {
      alert(`Undo failed: ${e.message}`);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.chatHistory}>
        {messages.map(m => (
          <div key={m.id} className={`${styles.message} ${m.role === 'user' ? styles.userMessage : styles.aiMessage}`}>
            {m.content}
            {m.isApprovalPending && (
              <div className={styles.approvalBox}>
                <p><strong>Warning:</strong> This action may overwrite data. Approve?</p>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <Button size="small" appearance="primary" onClick={() => handleApprove(m.id, m.proposedChanges)}>Approve</Button>
                  <Button size="small" onClick={() => setMessages(prev => prev.map(msg => msg.id === m.id ? { ...msg, isApprovalPending: false, content: msg.content + ' (Rejected)' } : msg))}>Reject</Button>
                </div>
              </div>
            )}
          </div>
        ))}
        {isProcessing && <div className={`${styles.message} ${styles.aiMessage}`}>Thinking...</div>}
      </div>
      
      {canUndo && (
        <Button style={{ marginBottom: '8px' }} onClick={handleUndo}>Undo Last AI Action</Button>
      )}

      <div className={styles.inputArea}>
        <Input 
          style={{ flex: 1 }}
          value={input}
          onChange={(_, d) => setInput(d.value)}
          onKeyDown={(e) => e.key === 'Enter' && handleSend()}
          placeholder="Ask AI to format or edit..."
          disabled={isProcessing}
        />
        <Button appearance="primary" onClick={handleSend} disabled={isProcessing}>Send</Button>
      </div>
    </div>
  );
};

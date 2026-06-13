import type { AppSettings } from './storage';
import type { ExcelContextInfo } from './excelContext';

export interface AIResponse {
  actionType: 'clarification' | 'edit' | 'destructive_edit' | 'read_more_data';
  messageToUser: string;
  proposedChanges?: {
    rangeAddress: string;
    values?: any[][];
    format?: any;
  };
}

const buildSystemPrompt = (excelContext: ExcelContextInfo): string => {
  return `Role: You are an expert Excel AI Assistant.
Objective: Analyze the user's request and current workbook context to propose data edits, formatting, or analysis.

Details (Current Environment):
- Active Sheet: ${excelContext.activeSheetName}
- Selected Range: ${excelContext.selectionAddress}
- All Sheets Context: ${JSON.stringify(excelContext.sheets.map(s => ({ sheet: s.sheetName, populatedRange: s.address, totalRows: s.totalRows, sample: s.sampleValues })))}

Rules:
1. If you need more data from a specific sheet/range before taking action, set actionType to "read_more_data" and specify the exact "rangeAddress" (e.g. "Sheet1!A1:Z100"). The add-in will fetch this data and reply to you automatically.
2. If the user's request is ambiguous, reply with a clarifying question. Set actionType to "clarification".
3. If the request involves destructive actions, set actionType to "destructive_edit".
4. For normal edits, set actionType to "edit" and provide the exact "rangeAddress" indicating where to place the "values". Always include the sheet name in the address (e.g. "Sheet2!A1"). Note: You only need to provide the top-left cell; the add-in will automatically resize the range to perfectly fit the dimensions of your provided "values" array. DO NOT use markdown like **bold** in the cell values! To format cells, use the "format" object.
5. The "format" object supports: { "bold": true, "italic": true, "backgroundColor": "#FFFF00", "fontColor": "#FF0000", "horizontalAlignment": "Center" | "Left" | "Right", "numberFormat": "0.00%" | "$#,##0.00" | etc, "convertToTable": true }. Only include properties you want to change. If convertToTable is true, the range will be converted to a native Excel Table.
6. Output must be valid JSON: { "actionType": "clarification" | "edit" | "destructive_edit" | "read_more_data", "messageToUser": "string", "proposedChanges": { "rangeAddress": "string", "values": [[]], "format": {} } }`;
};

const extractJSON = (text: string) => {
  try {
    return JSON.parse(text);
  } catch (e) {
    const match = text.match(/```(?:json)?\n?([\s\S]*?)\n?```/i);
    if (match) {
      try { return JSON.parse(match[1]); } catch(err) {}
    }
    
    // Attempt to extract raw JSON block if no markdown was used but there is text around it
    const firstBrace = text.indexOf('{');
    const lastBrace = text.lastIndexOf('}');
    if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
      try {
        return JSON.parse(text.substring(firstBrace, lastBrace + 1));
      } catch(err) {}
    }
    
    throw new Error("Failed to parse JSON from AI response: " + text);
  }
};

export const sendToAI = async (
  prompt: string,
  settings: AppSettings,
  excelContext: ExcelContextInfo,
  conversationHistory: {role: string, content: string}[] = []
): Promise<AIResponse> => {
  const systemPrompt = buildSystemPrompt(excelContext);
  
  const enforcedPrompt = `${prompt}\n\n[CRITICAL REMINDER: You MUST output ONLY valid JSON matching the required schema. Do NOT wrap it in markdown backticks, and do NOT include any conversational text before or after the JSON.]`;

  const messages = [
    { role: 'system', content: systemPrompt },
    ...conversationHistory,
    { role: 'user', content: enforcedPrompt }
  ];

  let endpoint = '';
  let headers: any = { 'Content-Type': 'application/json' };
  let body: any = {
    messages,
    temperature: 0.2,
    max_tokens: 4096
  };

  if (settings.provider === 'lmstudio') {
    endpoint = `${settings.lmStudioUrl}/chat/completions`;
    body.model = settings.lmStudioModel || "local-model";
  } else if (settings.provider === 'gemini') {
    endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent?key=${settings.geminiApiKey}`;
    body = {
      system_instruction: { parts: { text: systemPrompt } },
      contents: [
        ...conversationHistory.map(m => ({ role: m.role === 'system' ? 'user' : m.role, parts: [{ text: m.content }] })),
        { role: 'user', parts: [{ text: enforcedPrompt }] }
      ],
      generationConfig: { response_mime_type: "application/json" }
    };
  } else if (settings.provider === 'opencode') {
    endpoint = 'https://zen.opencode.ai/v1/chat/completions';
    headers['Authorization'] = `Bearer ${settings.opencodeApiKey}`;
    body.model = 'minimax-m2.5-free';
  }

  const res = await fetch(endpoint, {
    method: 'POST',
    headers,
    body: JSON.stringify(body)
  });

  if (!res.ok) {
    let errorDetail = res.statusText;
    try {
      const errorData = await res.json();
      errorDetail = JSON.stringify(errorData);
    } catch {
      try {
        const errorText = await res.text();
        if (errorText) errorDetail = errorText;
      } catch {}
    }
    throw new Error(`API Error (${res.status}): ${errorDetail}`);
  }

  const data = await res.json();
  
  if (settings.provider === 'gemini') {
    return extractJSON(data.candidates[0].content.parts[0].text);
  }

  return extractJSON(data.choices[0].message.content);
};

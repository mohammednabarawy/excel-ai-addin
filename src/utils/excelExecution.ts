// @ts-nocheck
export interface AIProposedChanges {
  rangeAddress: string;
  values?: any[][];
  format?: any; // e.g., { backgroundColor: 'yellow', bold: true }
}

let snapshot: {
  address: string;
  values: any[][];
  format: any;
  sheetName: string;
  tableCreated?: string;
} | null = null;

export const executeChange = async (changes: AIProposedChanges) => {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    let range = changes.rangeAddress ? sheet.getRange(changes.rangeAddress) : context.workbook.getSelectedRange();
    
    if (changes.values && changes.values.length > 0) {
      const rows = changes.values.length;
      const cols = changes.values[0].length;
      range = range.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    }
    
    // 1. Take Snapshot
    sheet.load("name");
    range.load("address, values, format/fill/color, format/font/bold, format/font/italic, format/font/color, format/horizontalAlignment, numberFormat");
    await context.sync();
    
    const actualAddress = changes.rangeAddress || range.address;

    snapshot = {
      address: actualAddress,
      values: range.values,
      sheetName: sheet.name,
      format: {
        backgroundColor: range.format.fill.color,
        bold: range.format.font.bold,
        italic: range.format.font.italic,
        fontColor: range.format.font.color,
        horizontalAlignment: range.format.horizontalAlignment,
        numberFormat: range.numberFormat
      }
    };

    // 2. Apply Changes
    if (changes.values && changes.values.length > 0) {
      range.values = changes.values;
    }
    
    if (changes.format) {
      if (changes.format.backgroundColor) {
        range.format.fill.color = changes.format.backgroundColor;
      }
      if (changes.format.bold !== undefined) {
        range.format.font.bold = changes.format.bold;
      }
      if (changes.format.italic !== undefined) {
        range.format.font.italic = changes.format.italic;
      }
      if (changes.format.fontColor) {
        range.format.font.color = changes.format.fontColor;
      }
      if (changes.format.horizontalAlignment) {
        range.format.horizontalAlignment = changes.format.horizontalAlignment;
      }
      if (changes.format.numberFormat) {
        range.numberFormat = changes.format.numberFormat;
      }
      if (changes.format.convertToTable) {
        const newTable = sheet.tables.add(range, true);
        newTable.load("name");
        await context.sync();
        snapshot.tableCreated = newTable.name;
      }
    }

    await context.sync();
  });
};

export const undoLastChange = async () => {
  if (!snapshot) throw new Error("No snapshot available to undo.");
  
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(snapshot.sheetName);
    const range = sheet.getRange(snapshot.address);
    
    range.values = snapshot.values;
    range.format.fill.color = snapshot.format.backgroundColor;
    range.format.font.bold = snapshot.format.bold;
    range.format.font.italic = snapshot.format.italic;
    range.format.font.color = snapshot.format.fontColor;
    range.format.horizontalAlignment = snapshot.format.horizontalAlignment;
    range.numberFormat = snapshot.format.numberFormat;

    if (snapshot.tableCreated) {
      const table = sheet.tables.getItemOrNullObject(snapshot.tableCreated);
      table.load("isNullObject");
      await context.sync();
      if (!table.isNullObject) {
        table.convertToRange();
      }
    }

    await context.sync();
    snapshot = null;
  });
};

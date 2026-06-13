// @ts-nocheck
export interface SheetSample {
  sheetName: string;
  address: string;
  sampleValues: any[][];
  totalRows: number;
  totalCols: number;
}

export interface ExcelContextInfo {
  activeSheetName: string;
  sheets: SheetSample[];
  selectionAddress: string;
  selectionValues: any[][];
}

export const getExcelContext = async (): Promise<ExcelContextInfo> => {
  return new Promise((resolve, reject) => {
    try {
      Excel.run(async (context) => {
        const workbook = context.workbook;
        const sheets = workbook.worksheets;
        const activeSheet = workbook.worksheets.getActiveWorksheet();
        const range = workbook.getSelectedRange();

        sheets.load("items/name");
        activeSheet.load("name");
        range.load("address, values");

        await context.sync();

        const sheetSamples: SheetSample[] = [];
        
        for (const sheet of sheets.items) {
          const usedRange = sheet.getUsedRangeOrNullObject();
          usedRange.load("address, rowCount, columnCount");
          await context.sync();

          if (!usedRange.isNullObject) {
            // Get up to first 5 rows for sample
            const rowsToFetch = Math.min(usedRange.rowCount, 5);
            const colsToFetch = Math.min(usedRange.columnCount, 20); // Limit columns to 20 to prevent huge JSONs
            const sampleRange = usedRange.getCell(0,0).getBoundingRect(
              usedRange.getCell(rowsToFetch - 1, colsToFetch - 1)
            );
            sampleRange.load("values");
            await context.sync();

            sheetSamples.push({
              sheetName: sheet.name,
              address: usedRange.address, // full used range
              sampleValues: sampleRange.values,
              totalRows: usedRange.rowCount,
              totalCols: usedRange.columnCount
            });
          } else {
             sheetSamples.push({
              sheetName: sheet.name,
              address: "",
              sampleValues: [],
              totalRows: 0,
              totalCols: 0
            });
          }
        }

        resolve({
          activeSheetName: activeSheet.name,
          sheets: sheetSamples,
          selectionAddress: range.address,
          selectionValues: range.values
        });
      }).catch(reject);
    } catch (e) {
      reject(e);
    }
  });
};

export const getSpecificRangeData = async (address: string): Promise<any[][]> => {
  return new Promise((resolve, reject) => {
    try {
      Excel.run(async (context) => {
        let range;
        if (address.includes('!')) {
          const [sheetName, rangeAddress] = address.split('!');
          // Remove single quotes if sheet name is wrapped in them (e.g., 'Sheet 1'!A1)
          const cleanSheetName = sheetName.replace(/^'|'$/g, '');
          range = context.workbook.worksheets.getItem(cleanSheetName).getRange(rangeAddress);
        } else {
          range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
        }
        range.load("values");
        await context.sync();
        resolve(range.values);
      }).catch(reject);
    } catch (e) {
      reject(e);
    }
  });
};

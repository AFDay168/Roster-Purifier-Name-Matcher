
import * as XLSX from 'xlsx';
import { RosterSheet } from '../types';

/**
 * Robustly parses a date string in yyyy/mm/dd format.
 * This avoids environment-specific interpretation of month and day order.
 */
const parseYMD = (val: any): Date | null => {
  if (val instanceof Date) return val;
  if (!val) return null;
  
  const str = String(val).trim();
  // Match yyyy/mm/dd or yyyy-mm-dd
  const parts = str.split(/[\/\-]/);
  if (parts.length === 3) {
    const y = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1; // 0-indexed
    const d = parseInt(parts[2], 10);
    
    // Create date using local time to avoid timezone shifts during component life
    const date = new Date(y, m, d, 0, 0, 0);
    if (!isNaN(date.getTime())) return date;
  }
  
  // Fallback to native parser if split fails
  const fallback = new Date(str);
  return isNaN(fallback.getTime()) ? null : fallback;
};

/**
 * Parses an Excel or CSV file into a collection of sheets with raw data.
 * Filters for tabs named strictly in yyyymmdd format for rosters.
 */
export const parseExcel = async (file: File): Promise<RosterSheet[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        
        const isCsv = file.name.toLowerCase().endsWith('.csv');
        const rosterTabPattern = /^\d{8}$/;
        const hasRosterTabs = workbook.SheetNames.some(name => rosterTabPattern.test(name.trim()));
        
        let targetSheetNames = workbook.SheetNames;
        if (hasRosterTabs && !isCsv) {
          targetSheetNames = workbook.SheetNames.filter(name => rosterTabPattern.test(name.trim()));
        }

        const sheets: RosterSheet[] = targetSheetNames.map((name) => {
          const worksheet = workbook.Sheets[name];
          // We use raw: false to get the displayed string format from the original file (yyyy/mm/dd)
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: null }) as any[][];
          return { name, data: jsonData };
        });
        resolve(sheets);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

/**
 * Extracts the majority month from Column C (index 2) across all sheets.
 */
export const findMajorityMonth = (sheets: RosterSheet[]): string | null => {
  const monthCounts: Record<string, number> = {};

  sheets.forEach((sheet) => {
    sheet.data.forEach((row, index) => {
      if (index === 0) return; // Skip headers
      const d = parseYMD(row[2]);
      if (d) {
        const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        monthCounts[monthKey] = (monthCounts[monthKey] || 0) + 1;
      }
    });
  });

  let maxCount = 0;
  let majority: string | null = null;

  Object.entries(monthCounts).forEach(([month, count]) => {
    if (count > maxCount) {
      maxCount = count;
      majority = month;
    }
  });

  return majority;
};

/**
 * Main cleaning logic:
 * 1. Filter by majority month in Column C.
 * 2. Keep columns A-H (indices 0-7).
 * 3. Keep rows 1-72 (indices 0-71).
 */
export const cleanRosterData = (sheets: RosterSheet[], majorityMonth: string): RosterSheet[] => {
  const processedSheets = sheets.map((sheet) => {
    let truncatedRows = sheet.data.slice(0, 72);

    const cleanedRows = truncatedRows.filter((row, index) => {
      if (index === 0) return true; // Keep headers

      const d = parseYMD(row[2]);
      if (!d) return false;

      const rowMonth = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      return rowMonth === majorityMonth;
    });

    const finalizedRows = cleanedRows.map((row, rIdx) => {
      const truncated = row.slice(0, 8);
      while (truncated.length < 8) truncated.push(null);
      
      if (rIdx > 0 && truncated[2]) {
        const d = parseYMD(truncated[2]);
        if (d) {
          truncated[2] = d;
        }
      }
      
      return truncated;
    });

    return {
      ...sheet,
      data: finalizedRows,
    };
  });

  return processedSheets.filter(sheet => sheet.data.length > 1);
};

/**
 * Updates names in Column F (index 5) of the Roster.
 */
export const updateNames = (rosterSheets: RosterSheet[], staffList: any[][]): RosterSheet[] => {
  const staffFullNames = staffList
    .map(row => row && row[0])
    .filter(name => name !== null && name !== undefined)
    .map(name => String(name).trim())
    .filter(name => name.length > 0);

  return rosterSheets.map((sheet) => ({
    ...sheet,
    data: sheet.data.map((row, rIdx) => {
      if (rIdx === 0 || !row) return row;

      const newRow = [...row];
      const currentRosterName = newRow[5];

      if (currentRosterName !== null && currentRosterName !== undefined) {
        let rosterNameStr = String(currentRosterName).trim();
        rosterNameStr = rosterNameStr.replace(/\s*\(.*?\)\s*/g, ' ').trim();
        const rosterNameLower = rosterNameStr.toLowerCase();

        if (rosterNameStr) {
          if (rosterNameLower === 'clara ckm') {
            const claraMatch = staffFullNames.find(sn => sn.toLowerCase().includes('clara cheung ka man'));
            if (claraMatch) {
              newRow[5] = claraMatch;
              return newRow;
            }
          }

          if (rosterNameLower === 'clara cheung') {
            const claraWingMatch = staffFullNames.find(sn => sn.toLowerCase().includes('clara cheung wing kum'));
            if (claraWingMatch) {
              newRow[5] = claraWingMatch;
              return newRow;
            }
          }

          const exactMatch = staffFullNames.find(sn => sn.toLowerCase() === rosterNameLower);
          if (exactMatch) {
            newRow[5] = exactMatch;
          } else {
            const partialMatch = staffFullNames.find(sn => 
              sn.toLowerCase().includes(rosterNameLower)
            );
            if (partialMatch) {
              newRow[5] = partialMatch;
            }
          }
        }
      }
      
      return newRow;
    }),
  }));
};

/**
 * Exports processed data back to an Excel file.
 * Ensures dates in Column C are saved as numeric date values with yyyy-mm-dd format.
 */
export const exportToExcel = (sheets: RosterSheet[], fileName: string) => {
  const wb = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    const numericData = sheet.data.map((row, rIdx) => {
      if (rIdx === 0) return row;
      const newRow = [...row];
      if (newRow[2]) {
        // Force re-parsing to ensure manual UI edits are also converted correctly
        const d = parseYMD(newRow[2]);
        if (d) {
          newRow[2] = d;
        }
      }
      return newRow;
    });

    const ws = XLSX.utils.aoa_to_sheet(numericData, { cellDates: true });
    
    if (ws['!ref']) {
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let r = range.s.r + 1; r <= range.e.r; ++r) {
        const cellRef = XLSX.utils.encode_cell({ r, c: 2 });
        if (ws[cellRef] && ws[cellRef].t === 'd') {
          // Explicitly set the format to ensure month and day are correctly presented
          ws[cellRef].z = 'yyyy-mm-dd';
        }
      }
    }

    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  });
  XLSX.writeFile(wb, fileName);
};

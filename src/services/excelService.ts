/* global Excel */

export async function executeInsertTable(headers: string[], records: any[]) {
  await Excel.run(async (context) => {
    const rows = records.map((r: any) => headers.map((h: string) => r[h] !== undefined ? r[h] : ""));
    const matrix = [headers, ...rows];
    const activeCell = context.workbook.getActiveCell();
    const range = activeCell.getResizedRange(matrix.length - 1, matrix[0].length - 1);
    range.values = matrix;
    range.format.wrapText = false;
    
    const table = context.workbook.tables.add(range, true);
    table.name = `DCTable_${Date.now()}`;
    table.style = "TableStyleLight1";
    table.showBandedRows = false;
    
    await context.sync();
  });
}

export async function executeInsertDropdown(headersList: string[]) {
  await Excel.run(async (context) => {
    let sourceString = headersList.join(",");
    if (sourceString.length >= 255) {
       let validHeaders: string[] = [];
       let currentLength = 0;
       for (let h of headersList) {
           if (currentLength + h.length + 1 < 255) {
               validHeaders.push(h);
               currentLength += h.length + 1;
           } else break;
       }
       headersList = validHeaders;
    }

    const range = context.workbook.getActiveCell();
    range.dataValidation.rule = {
      list: {
        inCellDropDown: true,
        source: headersList.join(",")
      }
    };
    await context.sync();
  });
}

export async function executeInsertFormula(id: string, field: string, dataTableName: string, selectedRev: number, isLatest: boolean) {
  await Excel.run(async (context) => {
    const revArg = isLatest ? `` : `, "${dataTableName}", ${selectedRev}`;
    const range = context.workbook.getActiveCell();
    const fieldArg = field ? `"${field}"` : `""`;
    range.formulas = [[`=DC.GET("${id}", ${fieldArg}${revArg ? revArg : `, "${dataTableName}"`})`]];
    await context.sync();
  });
}

export async function executeConvertToValues(): Promise<number> {
  let count = 0;
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();
    for (let sheet of sheets.items) {
      const usedRange = sheet.getUsedRangeOrNullObject();
      await context.sync();
      if (usedRange.isNullObject) continue;
      const formulaCells = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
      formulaCells.load("areas");
      await context.sync();
      if (formulaCells.isNullObject) continue;
      for (let area of formulaCells.areas.items) {
        area.load("formulas");
        await context.sync();
        const formulas = area.formulas;
        const dcCells = [];
        for (let i = 0; i < formulas.length; i++) {
          for (let j = 0; j < formulas[i].length; j++) {
            const cellFormula = formulas[i][j];
            if (typeof cellFormula === "string" && cellFormula.toUpperCase().includes("DC.")) {
              const cell = area.getCell(i, j);
              const spillRange = cell.getSpillingToRangeOrNullObject();
              cell.load("values");
              spillRange.load("values");
              dcCells.push({ cell, spillRange });
              count++;
            }
          }
        }
        if (dcCells.length > 0) {
          await context.sync();
          for (const item of dcCells) {
            if (!item.spillRange.isNullObject) {
              item.spillRange.values = item.spillRange.values; 
            } else {
              item.cell.values = item.cell.values; 
            }
          }
        }
      }
    }
    await context.sync();
  });
  return count;
}

export async function executeRefreshFormulas(): Promise<number> {
  let count = 0;
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();
    for (let sheet of sheets.items) {
      const usedRange = sheet.getUsedRangeOrNullObject();
      await context.sync();
      if (usedRange.isNullObject) continue;
      const formulaCells = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
      formulaCells.load("areas");
      await context.sync();
      if (formulaCells.isNullObject) continue;
      for (let area of formulaCells.areas.items) {
        area.load("formulas");
        await context.sync();
        const formulas = area.formulas;
        for (let i = 0; i < formulas.length; i++) {
          for (let j = 0; j < formulas[i].length; j++) {
            const cellFormula = formulas[i][j];
            if (typeof cellFormula === "string" && cellFormula.toUpperCase().includes("DC.")) {
              const cell = area.getCell(i, j);
              // Toggle an invisible space at the end to trick Excel into forcing a recalculation
              const newFormula = cellFormula.endsWith(" ") ? cellFormula.trim() : cellFormula + " ";
              cell.formulas = [[newFormula]]; 
              count++;
            }
          }
        }
      }
    }
    await context.sync();
  });
  return count;
}
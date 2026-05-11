/* global Office */
import { idbGet } from "../utils/db";

export async function getTableData(dataTableName?: string, rev?: number): Promise<{ records: any[], fields: string[], error?: string }> {
  try {
    await Office.onReady(); // Ensure Shared Runtime is fully initialized

    const defaultDataTable = await idbGet("DC_DEFAULT_DATA_TABLE");
    const targetDataTable = dataTableName && String(dataTableName).trim() !== "" ? String(dataTableName) : defaultDataTable;

    let finalRev = rev;
    if ((rev === undefined || rev === null) && targetDataTable === defaultDataTable) {
      const defRevStr = await idbGet("DC_DEFAULT_REVISION");
      if (defRevStr) finalRev = parseInt(defRevStr);
    }

    if (!targetDataTable) return { records: [], fields: [], error: "Error: No data table specified." };

    const storedData = await idbGet("DC_STORE");
    if (!storedData) return { records: [], fields: [], error: "Error: No data captured." };

    const store = JSON.parse(storedData);
    const dataSet = store[targetDataTable];
    if (!dataSet) return { records: [], fields: [], error: "Error: Data table not found." };

    let records = dataSet.records;
    let fields = dataSet.fields;

    if (finalRev !== undefined && finalRev !== null && finalRev !== (dataSet.revision || 1)) {
      if (dataSet.history && dataSet.history[finalRev]) {
        records = dataSet.history[finalRev].records;
        fields = dataSet.history[finalRev].fields;
      } else {
        return { records: [], fields: [], error: "Error: Revision not found." };
      }
    }

    return { records, fields };
  } catch (error) {
    return { records: [], fields: [], error: `DB Error: ${error.message || error}` };
  }
}

export async function getGlobalVariable(varName: string): Promise<any> {
  try {
      await Office.onReady();
      const vStoreRaw = await idbGet("DC_VARIABLES");
      const variables = vStoreRaw ? JSON.parse(vStoreRaw) : {};
      const formula = variables[varName];
      if (!formula) return "Error: Variable not found";

      const storedData = await idbGet("DC_STORE");
      const store = storedData ? JSON.parse(storedData) : {};
      
      const evaluateVar = (vName: string, visited: Set<string>): any => {
          if (visited.has(vName)) throw new Error("Circular reference detected");
          visited.add(vName);
          const vForm = variables[vName];
          if (!vForm) return 0;
          const DC = {
              SUM: (subTable: string, sumCol: string, fkCol?: string, fkValue?: any) => {
                  if (!store || !store[subTable]) return 0;
                  let total = 0;
                  store[subTable].records.forEach((r: any) => { if (!fkCol || String(r[fkCol]) === String(fkValue)) total += Number(r[sumCol]) || 0; });
                  return total;
              },
              COUNT: (subTable: string, fkCol?: string, fkValue?: any) => {
                  if (!store || !store[subTable]) return 0;
                  let count = 0;
                  store[subTable].records.forEach((r: any) => { if (!fkCol || String(r[fkCol]) === String(fkValue)) count++; });
                  return count;
              },
              VAR: (v: string) => evaluateVar(v, new Set(visited))
          };
          const func = new Function('store', 'DC', `return ${vForm};`);
          return func(store, DC);
      };

      return evaluateVar(varName, new Set());
  } catch (error) {
      return `Error: ${error.message || error}`;
  }
}
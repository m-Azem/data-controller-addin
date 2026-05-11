﻿import { getTableData, getGlobalVariable } from "../services/dataService";
import { idbGet } from "../utils/db";

/**
 * Retrieves a value from the captured Data Controller store by matching the Record ID.
 * @customfunction GET
 * @param {any} [id] Optional. The unique Record ID to find. Returns headers if empty.
 * @param {string} [fieldName] Optional. The name of the property to retrieve. Returns the entire record if empty.
 * @param {string} [dataTableName] Optional. The name of the data table. Uses default if empty.
 * @param {number} [rev] Optional. The revision number. Uses the latest by default.
 * @returns The value from the data store.
 */
export async function dcGet(id?: any, fieldName?: string, dataTableName?: string, rev?: number): Promise<any[][]> {
  try {
    const table = await getTableData(dataTableName, rev);
    if (table.error) return [[table.error]];

    const { records, fields: headers } = table as any;
    
    const storeRaw = await idbGet("DC_STORE");
    const store = storeRaw ? JSON.parse(storeRaw) : {};
    const defaultTable = await idbGet("DC_DEFAULT_DATA_TABLE");
    const resolvedTableName = dataTableName && String(dataTableName).trim() !== "" ? dataTableName : defaultTable;
    const dataSet = store[resolvedTableName] || {};
    
    if (id === undefined || id === null || String(id).trim() === "") {
      return [headers];
    }

    const record = records.find((r: any) => String(r.__DC_ID__) === String(id));

    if (!record) {
      return [["Not Found"]];
    }
    
    if (!fieldName || String(fieldName).trim() === "") {
      return [headers.map((h: string) => record[h] !== undefined ? record[h] : "")];
    }
    return [[record[fieldName] !== undefined ? record[fieldName] : "Field Error"]];
  } catch (error) {
    return [[`Func Error: ${error.message || error}`]];
  }
}

/**
 * Retrieves the evaluated result of a Global Variable.
 * @customfunction VAR
 * @param {string} varName The exact name of the Global Variable to retrieve.
 * @returns The evaluated variable result.
 */
export async function dcVar(varName: string): Promise<any[][]> {
  try {
    const result = await getGlobalVariable(varName);
    return [[result]];
  } catch (error: any) {
    return [[`Func Error: ${error.message || error}`]];
  }
}

/**
 * Searches for a value in a specific column and returns multiple matching records.
 * @customfunction SEARCH
 * @param {string} searchField The column name to search in.
 * @param {any} searchValue The value to find.
 * @param {string} returnField The column name of the value to return.
 * @param {string} [dataTableName] Optional. The name of the data table. Uses default if empty.
 * @param {number} [rev] Optional. The revision number. Uses the latest by default.
 * @returns A dynamic array of matching values.
 */
export async function dcSearch(searchField: string, searchValue: any, returnField: string, dataTableName?: string, rev?: number): Promise<any[][]> {
  try {
    const table = await getTableData(dataTableName, rev);
    if (table.error) return [[table.error]];
    
    const searchStr = String(searchValue).toLowerCase();
    const results = table.records
      .filter((r: any) => r[searchField] !== undefined && String(r[searchField]).toLowerCase().includes(searchStr))
      .map((r: any) => [r[returnField] !== undefined ? r[returnField] : ""]);

    if (results.length === 0) return [["Not Found"]];
    return results;
  } catch (error) {
    return [[`Func Error: ${error.message || error}`]];
  }
}

/**
 * Returns all full records that match a specific criteria.
 * @customfunction FILTER
 * @param {string} searchField The column name to search in.
 * @param {any} searchValue The value to find.
 * @param {string} [dataTableName] Optional. The name of the data table.
 * @param {number} [rev] Optional. The revision number.
 * @param {boolean} [exactMatch] Optional. Set to true for an exact match. Default is false (partial match).
 * @returns A dynamic array of matching records (including headers).
 */
export async function dcFilter(searchField: string, searchValue: any, dataTableName?: string, rev?: number, exactMatch?: boolean): Promise<any[][]> {
  try {
    const table = await getTableData(dataTableName, rev);
    if (table.error) return [[table.error]];

    const searchStr = String(searchValue).toLowerCase();
    const matched = table.records.filter((r: any) => {
      if (r[searchField] === undefined) return false;
      const recordStr = String(r[searchField]).toLowerCase();
      return exactMatch ? recordStr === searchStr : recordStr.includes(searchStr);
    });

    if (matched.length === 0) return [["Not Found"]];

    const resultRows = matched.map((r: any) => table.fields.map(f => r[f] !== undefined ? r[f] : ""));
    return [table.fields, ...resultRows];
  } catch (error) {
    return [[`Func Error: ${error.message || error}`]];
  }
}

/**
 * Sums all numerical values in a specified column.
 * @customfunction SUM
 * @param {string} sumField The column name to sum.
 * @param {string} [dataTableName] Optional. The name of the data table.
 * @param {number} [rev] Optional. The revision number.
 * @returns The total sum.
 */
export async function dcSum(sumField: string, dataTableName?: string, rev?: number): Promise<number | string> {
  try {
    const table = await getTableData(dataTableName, rev);
    if (table.error) return table.error;

    let total = 0;
    for (const r of table.records) {
      const val = parseFloat(r[sumField]);
      if (!isNaN(val)) total += val;
    }
    return total;
  } catch (error) {
    return `Func Error: ${error.message || error}`;
  }
}

/**
 * Sums values in a column based on a specific condition.
 * @customfunction SUMIFS
 * @param {string} sumField The column name to sum.
 * @param {string} criteriaField The column name to check.
 * @param {any} criteriaValue The value to match.
 * @param {string} [dataTableName] Optional. The name of the data table.
 * @param {number} [rev] Optional. The revision number.
 * @returns The total sum.
 */
export async function dcSumifs(sumField: string, criteriaField: string, criteriaValue: any, dataTableName?: string, rev?: number): Promise<number | string> {
  try {
    const table = await getTableData(dataTableName, rev);
    if (table.error) return table.error;

    let total = 0;
    const searchStr = String(criteriaValue).toLowerCase();
    for (const r of table.records) {
      if (r[criteriaField] !== undefined && String(r[criteriaField]).toLowerCase().includes(searchStr)) {
        const val = parseFloat(r[sumField]);
        if (!isNaN(val)) total += val;
      }
    }
    return total;
  } catch (error) {
    return `Func Error: ${error.message || error}`;
  }
}

/**
 * Joins a base table with a foreign table and appends a selected column from the foreign table.
 * @customfunction JOIN
 * @param {string} baseTableName The name of the primary data table (e.g., Orders).
 * @param {string} foreignKeyField The column in the base table that links to the foreign table (e.g., CustomerID).
 * @param {string} foreignTableName The name of the related table to look up (e.g., Customers).
 * @param {string} foreignReturnField The column in the foreign table to retrieve and append (e.g., CustomerName).
 * @returns A dynamic array of the joined table.
 */
export async function dcJoin(baseTableName: string, foreignKeyField: string, foreignTableName: string, foreignReturnField: string): Promise<any[][]> {
  try {
    const baseTable = await getTableData(baseTableName);
    if (baseTable.error) return [[`Base Table Error: ${baseTable.error}`]];

    const foreignTable = await getTableData(foreignTableName);
    if (foreignTable.error) return [[`Foreign Table Error: ${foreignTable.error}`]];

    const storeRaw = await idbGet("DC_STORE");
    const store = storeRaw ? JSON.parse(storeRaw) : {};
    const foreignDataSet = store[foreignTableName] || {};

    const newHeaders = [...baseTable.fields, foreignReturnField];

    const resultRows = baseTable.records.map((bRec: any) => {
      const fKey = bRec[foreignKeyField];
      const fRec = foreignTable.records.find((f: any) => String(f.__DC_ID__) === String(fKey));
      const fVal = fRec && fRec[foreignReturnField] !== undefined ? fRec[foreignReturnField] : "N/A";
      
      const rowData = baseTable.fields.map(f => bRec[f] !== undefined ? bRec[f] : "");
      return [...rowData, fVal];
    });

    return [newHeaders, ...resultRows];
  } catch (error) {
    return [[`Func Error: ${error.message || error}`]];
  }
}

/**
 * Returns all records sorted by a specific column.
 * @customfunction SORT
 * @param {string} sortField The column name to sort by.
 * @param {boolean} [ascending] Optional. Sort ascending (TRUE) or descending (FALSE). Default is TRUE.
 * @param {string} [dataTableName] Optional. The name of the data table. Uses default if empty.
 * @param {number} [rev] Optional. The revision number. Uses the latest by default.
 * @returns A dynamic array of sorted records (including headers).
 */
export async function dcSort(sortField: string, ascending: boolean = true, dataTableName?: string, rev?: number): Promise<any[][]> {
  try {
    const table = await getTableData(dataTableName, rev);
    if (table.error) return [[table.error]];

    if (!table.fields.includes(sortField)) {
      return [[`Sort Field '${sortField}' not found`]];
    }

    const sortedRecords = [...table.records].sort((a: any, b: any) => {
      let valA = a[sortField];
      let valB = b[sortField];
      
      if (valA === undefined || valA === null) valA = "";
      if (valB === undefined || valB === null) valB = "";

      const numA = Number(valA);
      const numB = Number(valB);
      
      if (!isNaN(numA) && !isNaN(numB)) {
        return ascending ? numA - numB : numB - numA;
      }

      const strA = String(valA).toLowerCase();
      const strB = String(valB).toLowerCase();

      if (strA < strB) return ascending ? -1 : 1;
      if (strA > strB) return ascending ? 1 : -1;
      return 0;
    });

    const resultRows = sortedRecords.map((r: any) => table.fields.map(f => r[f] !== undefined ? r[f] : ""));
    return [table.fields, ...resultRows];
  } catch (error) {
    return [[`Func Error: ${error.message || error}`]];
  }
}
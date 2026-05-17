﻿﻿﻿import { getTableData, getGlobalVariable } from "../services/dataService";
import { idbGet } from "../utils/db";
import { t } from "../taskpane/taskpane";

const requestCache = new Map<string, Promise<any>>();
const recordIdIndex = new WeakMap<any[], Map<string, any>>();

function withCache<T>(key: string, fn: () => Promise<T>): Promise<T> {
  if (requestCache.has(key)) return requestCache.get(key) as Promise<T>;
  const promise = fn();
  requestCache.set(key, promise);
  promise.finally(() => setTimeout(() => requestCache.delete(key), 10));
  return promise;
}

function getRecordById(records: any[], id: string) {
  let indexMap = recordIdIndex.get(records);
  if (!indexMap) {
    indexMap = new Map();
    for (let i = 0; i < records.length; i++) {
      if (records[i] && records[i].__DC_ID__ !== undefined) {
        indexMap.set(String(records[i].__DC_ID__), records[i]);
      }
    }
    recordIdIndex.set(records, indexMap);
  }
  return indexMap.get(id);
}

/**
 * Retrieves a value from the captured Data Controller store by matching the Record ID.
 * @customfunction GET
 * @param {any} [id] Optional. The unique Record ID to find. Returns headers if empty.
 * @param {string} [fieldName] Optional. The name of the property to retrieve. Returns the entire record if empty.
 * @param {string} [dataTableName] Optional. The name of the data table. Uses default if empty.
 * @param {number} [rev] Optional. The revision number. Uses the latest by default.
 * @returns The value from the data store.
 */
export function dcGet(id?: any, fieldName?: string, dataTableName?: string, rev?: number): Promise<any[][]> {
  const cacheKey = `GET|${id}|${fieldName}|${dataTableName}|${rev}`;
  return withCache(cacheKey, async () => {
    try {
      const table = await getTableData(dataTableName, rev);
      if (table.error) return [[table.error]];

      const { records, fields: headers } = table as any;
      
      if (id === undefined || id === null || String(id).trim() === "") {
        return [headers];
      }

      const record = getRecordById(records, String(id));

      if (!record) {
        return [[t("not_found")]];
      }
      
      if (!fieldName || String(fieldName).trim() === "") {
        return [headers.map((h: string) => record[h] !== undefined ? record[h] : "")];
      }
      return [[record[fieldName] !== undefined ? record[fieldName] : t("field_error")]];
    } catch (error: any) {
      return [[`${t("func_error_prefix")}${error.message || error}`]];
    }
  });
}

/**
 * Retrieves the evaluated result of a Global Variable.
 * @customfunction VAR
 * @param {string} varName The exact name of the Global Variable to retrieve.
 * @returns The evaluated variable result.
 */
export function dcVar(varName: string): Promise<any[][]> {
  const cacheKey = `VAR|${varName}`;
  return withCache(cacheKey, async () => {
    try {
      const result = await getGlobalVariable(varName);
      return [[result]];
    } catch (error: any) {
      return [[`${t("func_error_prefix")}${error.message || error}`]];
    }
  });
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
export function dcSearch(searchField: string, searchValue: any, returnField: string, dataTableName?: string, rev?: number): Promise<any[][]> {
  const cacheKey = `SEARCH|${searchField}|${searchValue}|${returnField}|${dataTableName}|${rev}`;
  return withCache(cacheKey, async () => {
    try {
      const table = await getTableData(dataTableName, rev);
      if (table.error) return [[table.error]];
      
      const searchStr = String(searchValue).toLowerCase();
      const results = table.records
        .filter((r: any) => r[searchField] !== undefined && String(r[searchField]).toLowerCase().includes(searchStr))
        .map((r: any) => [r[returnField] !== undefined ? r[returnField] : ""]);

      if (results.length === 0) return [[t("not_found")]];
      return results;
    } catch (error: any) {
      return [[`${t("func_error_prefix")}${error.message || error}`]];
    }
  });
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
export function dcFilter(searchField: string, searchValue: any, dataTableName?: string, rev?: number, exactMatch?: boolean): Promise<any[][]> {
  const cacheKey = `FILTER|${searchField}|${searchValue}|${dataTableName}|${rev}|${exactMatch}`;
  return withCache(cacheKey, async () => {
    try {
      const table = await getTableData(dataTableName, rev);
      if (table.error) return [[table.error]];

      const searchStr = String(searchValue).toLowerCase();
      
      const matched = exactMatch
        ? table.records.filter((r: any) => r[searchField] !== undefined && String(r[searchField]).toLowerCase() === searchStr)
        : table.records.filter((r: any) => r[searchField] !== undefined && String(r[searchField]).toLowerCase().includes(searchStr));

      if (matched.length === 0) return [[t("not_found")]];

      const resultRows = matched.map((r: any) => table.fields.map(f => r[f] !== undefined ? r[f] : ""));
      return [table.fields, ...resultRows];
    } catch (error: any) {
      return [[`${t("func_error_prefix")}${error.message || error}`]];
    }
  });
}

/**
 * Sums all numerical values in a specified column.
 * @customfunction SUM
 * @param {string} sumField The column name to sum.
 * @param {string} [dataTableName] Optional. The name of the data table.
 * @param {number} [rev] Optional. The revision number.
 * @returns The total sum.
 */
export function dcSum(sumField: string, dataTableName?: string, rev?: number): Promise<number | string> {
  const cacheKey = `SUM|${sumField}|${dataTableName}|${rev}`;
  return withCache(cacheKey, async () => {
    try {
      const table = await getTableData(dataTableName, rev);
      if (table.error) return table.error;
      
      let total = 0;
      for (const r of table.records) {
        const val = parseFloat(r[sumField]);
        if (!isNaN(val)) total += val;
      }
      return total;
    } catch (error: any) {
      return `${t("func_error_prefix")}${error.message || error}`;
    }
  });
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
export function dcSumifs(sumField: string, criteriaField: string, criteriaValue: any, dataTableName?: string, rev?: number): Promise<number | string> {
  const cacheKey = `SUMIFS|${sumField}|${criteriaField}|${criteriaValue}|${dataTableName}|${rev}`;
  return withCache(cacheKey, async () => {
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
    } catch (error: any) {
      return `${t("func_error_prefix")}${error.message || error}`;
    }
  });
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
export function dcJoin(baseTableName: string, foreignKeyField: string, foreignTableName: string, foreignReturnField: string): Promise<any[][]> {
  const cacheKey = `JOIN|${baseTableName}|${foreignKeyField}|${foreignTableName}|${foreignReturnField}`;
  return withCache(cacheKey, async () => {
    try {
      const baseTable = await getTableData(baseTableName);
      if (baseTable.error) return [[`${t("base_table_error_prefix")}${baseTable.error}`]];

      const foreignTable = await getTableData(foreignTableName);
      if (foreignTable.error) return [[`${t("foreign_table_error_prefix")}${foreignTable.error}`]];

      const newHeaders = [...baseTable.fields, foreignReturnField];

      const resultRows = baseTable.records.map((bRec: any) => {
        const fKey = bRec[foreignKeyField];
        const fRec = getRecordById(foreignTable.records, String(fKey));
        const fVal = fRec && fRec[foreignReturnField] !== undefined ? fRec[foreignReturnField] : t("na_value");
        
        const rowData = baseTable.fields.map(f => bRec[f] !== undefined ? bRec[f] : "");
        return [...rowData, fVal];
      });

      return [newHeaders, ...resultRows];
    } catch (error: any) {
      return [[`${t("func_error_prefix")}${error.message || error}`]];
    }
  });
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
export function dcSort(sortField: string, ascending: boolean = true, dataTableName?: string, rev?: number): Promise<any[][]> {
  const cacheKey = `SORT|${sortField}|${ascending}|${dataTableName}|${rev}`;
  return withCache(cacheKey, async () => {
    try {
      const table = await getTableData(dataTableName, rev);
      if (table.error) return [[table.error]];

      if (!table.fields.includes(sortField)) {
        return [[t("sort_field_not_found", sortField)]];
      }

      const sortIndex = table.records.map((r: any, index: number) => {
        let val = r[sortField];
        if (val === undefined || val === null) val = "";
        const num = Number(val);
        const isNum = !isNaN(num) && val !== "";
        return { 
          index, 
          num, 
          isNum, 
          str: isNum ? "" : String(val).toLowerCase(),
          record: r 
        };
      });

      sortIndex.sort((a, b) => {
        if (a.isNum && b.isNum) {
          return ascending ? a.num - b.num : b.num - a.num;
        }
        if (a.str < b.str) return ascending ? -1 : 1;
        if (a.str > b.str) return ascending ? 1 : -1;
        return a.index - b.index;
      });

      const resultRows = sortIndex.map(item => table.fields.map(f => item.record[f] !== undefined ? item.record[f] : ""));
      return [table.fields, ...resultRows];
    } catch (error: any) {
      return [[`${t("func_error_prefix")}${error.message || error}`]];
    }
  });
}
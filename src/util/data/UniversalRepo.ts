/// <reference types="office-js" />

import { excelLog } from "../Logs";

/*
  Excel Workbook ORM (metadata + typed CRUD for Tables)
  ----------------------------------------------------
  Highlights
  - TableRepository.getAs<T>(name, def) → TypedTable<T>
  - TypedRowRepository: getAll, findFirstBy/findAllBy, add, setAll, updateBy, deleteBy
  - Honors column metadata from TableDefinition: required, default, final, calculated

  NOTE: For calculated columns we never write values — we let Excel formulas populate them.
        For final columns we allow values on INSERT, but forbid changes on UPDATE.
*/

// ========================= Shared Types =========================
export type ColumnType = "string" | "number" | "boolean" | "date" | "any";

export interface ColumnSchema {
  name: string;
  id?: string; // Office sometimes exposes column id
  index: number;
}

export interface TableSchema {
  id: string;
  name: string;
  worksheet: string; // Worksheet name
  columns: ColumnSchema[];
}

export type DefaultValue<T> = T | (() => T);

export type TableDefinition<T extends Record<string, any>> = {
  /** Column contracts (+ optional metadata) */
  columns: { [K in keyof T]-?: { type?: ColumnType; required?: boolean; default?: DefaultValue<T[K]>; final?: boolean; calculated?: boolean; referenceTo?: string;} };
  /** Optional: mapping from property key → Excel header name (if different) */
  names?: Partial<Record<keyof T, string>>;
  /** Optional: write order for columns; read always respects actual header order */
  order?: (keyof T)[];
};

/** Options for row searching */
export interface RowFindOptions { caseInsensitive?: boolean; trim?: boolean; }

/** Untyped row match result */
export interface RowMatch { index: number; row: Record<string, any>; range: Excel.Range; }

/** Typed row match result */
export type RowMatchTyped<T> = { index: number; row: T; range: Excel.Range };

function normalizeForCompare(v: any, opts?: RowFindOptions): any {
  if (v == null) return v;
  if (v instanceof Date) return v.getTime();
  if (typeof v === "string") {
    let s = v;
    if (opts?.trim) s = s.trim();
    if (opts?.caseInsensitive) s = s.toLowerCase();
    return s;
  }
  return v;
}

function isEqual(a: any, b: any, opts?: RowFindOptions): boolean {
  return normalizeForCompare(a, opts) === normalizeForCompare(b, opts);
}

function isBlank(v: any): boolean {
  return v === undefined || v === null || v === "";
}

/** Utility helpers for Office.js loading */
// In @types/office-js, `load` is not declared on the base ClientObject type, but on each concrete object.
// So we constrain to a ClientObject that also has a `load` method.

type OfficeLoadable = OfficeExtension.ClientObject & { load: (props?: any) => any };

async function loadAndSync<T extends OfficeLoadable>(
  obj: T,
  props: string | string[]
): Promise<void> {
  obj.load(props as any);
  await (obj.context as Excel.RequestContext).sync();
}

/** Convert a 2D any[][] range values + headers to array of objects */
function rowsToObjects(headers: string[], values: any[][]): Record<string, any>[] {
  return values.map((row) => {
    const rec: Record<string, any> = {};
    headers.forEach((h, i) => (rec[h] = row[i]));
    return rec;
  });
}

/** Convert object to row array using provided headers */
function objectToRow(headers: string[], obj: Record<string, any>): any[] {
  return headers.map((h) => obj[h]);
}

// ========================= ORM Root =========================
export class WorkbookORM {
  private workbook: Excel.Workbook;
  public readonly tables: TableRepository;

  constructor(workbook: Excel.Workbook) {
    this.workbook = workbook;
    this.tables = new TableRepository(workbook);
  }
}

// ========================= Table Repository =========================
export class TableRepository {
  private workbook: Excel.Workbook;
  private cacheByName: Map<string, ExcelTable> = new Map();

  constructor(workbook: Excel.Workbook) { this.workbook = workbook; }

  /** Returns lightweight schemas for all tables in the workbook */
  async listSchemas(): Promise<TableSchema[]> {
    const ctx = this.workbook.context as Excel.RequestContext;
    const tables = this.workbook.tables;
    tables.load(["items"]);
    await ctx.sync();

    const results: TableSchema[] = [];
    for (const t of tables.items) {
      t.load(["id", "name", "worksheet/name", "columns/items/name", "columns/items/index"]);
    }
    await ctx.sync();

    for (const t of tables.items) {
      const columns: ColumnSchema[] = t.columns.items.map((c) => ({ name: c.name, index: c.index }));
      results.push({ id: t.id, name: t.name, worksheet: t.worksheet.name, columns });
    }
    return results.sort((a, b) => a.name.localeCompare(b.name));
  }

  /** Get an untyped wrapped table by name (case-sensitive Excel name) */
  async get(name: string): Promise<ExcelTable> {
    const cached = this.cacheByName.get(name); if (cached) return cached;
    const ctx = this.workbook.context as Excel.RequestContext;
    const table = this.workbook.tables.getItem(name);
    table.load(["id", "name", "worksheet/name", "columns/items/name", "columns/items/index"]);
    await ctx.sync();
    const wrapped = new ExcelTable(table);
    this.cacheByName.set(name, wrapped);
    return wrapped;
  }

  /** Get a typed table wrapper */
  async getAs<T extends Record<string, any>>(name: string, def: TableDefinition<T>): Promise<TypedTable<T>> {
    const base = await this.get(name);
    return new TypedTable<T>(base, def);
  }
}

// ========================= Untyped Table =========================
export class ExcelTable {
  private table: Excel.Table;
  public readonly columns: ColumnRepository;
  public readonly rows: RowRepository;

  constructor(table: Excel.Table) {
    this.table = table;
    this.columns = new ColumnRepository(table);
    this.rows = new RowRepository(table);
  }

  /** Schema snapshot */
  async getSchema(): Promise<TableSchema> {
    const t = this.table;
    await loadAndSync(t, ["id", "name", "worksheet/name", "columns/items/name", "columns/items/index"]);
    const columns: ColumnSchema[] = t.columns.items.map((c) => ({ name: c.name, index: c.index }));
    return { id: t.id, name: t.name, worksheet: t.worksheet.name, columns };
  }

  /** Native reference */
  get native(): Excel.Table { return this.table; }
}

// ========================= Column Access (Untyped) =========================
export class ColumnRepository {
  private table: Excel.Table;
  private byName: Map<string, ExcelColumn> = new Map();
  constructor(table: Excel.Table) { this.table = table; }

  async list(): Promise<ExcelColumn[]> {
    const ctx = this.table.context as Excel.RequestContext;
    this.table.columns.load(["items/name", "items/index"]);
    await ctx.sync();
    return this.table.columns.items
      .sort((a, b) => a.index - b.index)
      .map((col) => new ExcelColumn(this.table, col.name, col.index));
  }

  async get(name: string): Promise<ExcelColumn> {
    const cached = this.byName.get(name); if (cached) return cached;
    const ctx = this.table.context as Excel.RequestContext;
    const col = this.table.columns.getItem(name);
    col.load(["name", "index"]);
    await ctx.sync();
    const wrapped = new ExcelColumn(this.table, col.name, col.index);
    this.byName.set(name, wrapped);
    return wrapped;
  }
}

export class ExcelColumn {
  private table: Excel.Table;
  readonly name: string;
  readonly index: number; // zero-based index within table
  constructor(table: Excel.Table, name: string, index: number) { this.table = table; this.name = name; this.index = index; }

  async getDataRange(): Promise<Excel.Range> {
    const col = this.table.columns.getItem(this.name);
    const rng = col.getDataBodyRange();
    await loadAndSync(rng, ["values", "rowCount", "columnCount", "address"]);
    return rng;
  }

  async getValues(): Promise<any[]> {
    const col = this.table.columns.getItem(this.name);
    const rng = col.getDataBodyRange();
    rng.load(["values"]);
    await (this.table.context as Excel.RequestContext).sync();
    return (rng.values || []).map((r) => r[0]);
  }

  async setValues(values: any[]): Promise<void> {
    const col = this.table.columns.getItem(this.name);
    const rng = col.getDataBodyRange();
    rng.load(["rowCount"]);
    await (this.table.context as Excel.RequestContext).sync();
    if ((rng as any).rowCount !== values.length) {
      throw new Error(`Length mismatch for column "${this.name}": table has ${(rng as any).rowCount} rows, received ${values.length}`);
    }
    rng.values = values.map((v) => [v]);
    await (this.table.context as Excel.RequestContext).sync();
  }
}

// ========================= Row Helpers (Untyped) =========================
export class RowRepository {
  private table: Excel.Table;
  constructor(table: Excel.Table) { this.table = table; }

  protected async headers(): Promise<string[]> {
    const headerRange = this.table.getHeaderRowRange();
    headerRange.load(["values"]);
    await (this.table.context as Excel.RequestContext).sync();
    const headers = (headerRange.values?.[0] || []) as string[];
    return headers.map((h) => String(h));
  }

  async getAll(): Promise<Record<string, any>[]> {
    const headers = await this.headers();
    const ctx = this.table.context as Excel.RequestContext;
    const body: any = (this.table as any).getDataBodyRangeOrNullObject ? (this.table as any).getDataBodyRangeOrNullObject() : this.table.getDataBodyRange();
    body.load(["values", "rowCount", "isNullObject"]);
    await ctx.sync();
    const isEmpty = Boolean(body.isNullObject) || (body.rowCount === 0);
    const values: any[][] = isEmpty ? [] : (body.values || []);
    return rowsToObjects(headers, values);
  }

  async add(obj: Record<string, any>): Promise<void> {
    const headers = await this.headers();
    const row = objectToRow(headers, obj);
    this.table.rows.add(null /* add to end */, [row]);
    await (this.table.context as Excel.RequestContext).sync();
  }

  async setAll(objs: Record<string, any>[]): Promise<void> {
    const headers = await this.headers();
    const rows = objs.map((o) => objectToRow(headers, o));
    const body = this.table.getDataBodyRange();
    body.load(["rowCount"]);
    await (this.table.context as Excel.RequestContext).sync();
    if ((body as any).rowCount > 0) {
      body.clear(Excel.ClearApplyTo.contents);
      await (this.table.context as Excel.RequestContext).sync();
    }
    if (rows.length > 0) {
      this.table.rows.add(null, rows);
      await (this.table.context as Excel.RequestContext).sync();
    }
  }

  async clear(): Promise<void> {
    const body = this.table.getDataBodyRange();
    body.load(["rowCount"]);
    await (this.table.context as Excel.RequestContext).sync();
    if ((body as any).rowCount > 0) {
      body.clear(Excel.ClearApplyTo.contents);
      await (this.table.context as Excel.RequestContext).sync();
    }
  }

  /** Find first row by header name and value */
  async findFirstByHeader(header: string, value: any, opts?: RowFindOptions): Promise<RowMatch | null> {
    const headers = await this.headers();
    const colIdx = headers.indexOf(header);
    if (colIdx === -1) throw new Error(`Header not found: ${header}`);

    const column: any = this.table.columns.getItem(header);
    const colRange: any = column.getDataBodyRangeOrNullObject ? column.getDataBodyRangeOrNullObject() : column.getDataBodyRange();
    colRange.load(["values", "rowCount", "isNullObject"]);
    await (this.table.context as Excel.RequestContext).sync();
    if (Boolean(colRange.isNullObject) || colRange.rowCount === 0) return null;

    const colValues: any[] = (colRange.values || []).map((r: any[]) => r[0]);
    const foundIndex = colValues.findIndex((v) => isEqual(v, value, opts));
    if (foundIndex === -1) return null;

    const body: any = (this.table as any).getDataBodyRangeOrNullObject ? (this.table as any).getDataBodyRangeOrNullObject() : this.table.getDataBodyRange();
    body.load(["values", "isNullObject"]);
    await (this.table.context as Excel.RequestContext).sync();
    if (Boolean(body.isNullObject)) return null;
    const rowVals = (body.values || [])[foundIndex];
    const rowObj = rowsToObjects(headers, [rowVals])[0];

    const range = this.table.rows.getItemAt(foundIndex).getRange();
    return { index: foundIndex, row: rowObj, range };
  }

  /** Find all rows by header name and value */
  async findAllByHeader(header: string, value: any, opts?: RowFindOptions): Promise<RowMatch[]> {
    const headers = await this.headers();
    const colIdx = headers.indexOf(header);
    if (colIdx === -1) throw new Error(`Header not found: ${header}`);

    const column: any = this.table.columns.getItem(header);
    const colRange: any = column.getDataBodyRangeOrNullObject ? column.getDataBodyRangeOrNullObject() : column.getDataBodyRange();
    colRange.load(["values", "rowCount", "isNullObject"]);
    await (this.table.context as Excel.RequestContext).sync();

    const colValues: any[] = (colRange.values || []).map((r: any[]) => r[0]);

    const body: any = (this.table as any).getDataBodyRangeOrNullObject ? (this.table as any).getDataBodyRangeOrNullObject() : this.table.getDataBodyRange();
    body.load(["values", "isNullObject"]);
    await (this.table.context as Excel.RequestContext).sync();

    const matches: RowMatch[] = [];
    colValues.forEach((v, i) => {
      if (isEqual(v, value, opts)) {
        const rowVals = (body.values || [])[i];
        const rowObj = rowsToObjects(headers, [rowVals])[0];
        const range = this.table.rows.getItemAt(i).getRange();
        matches.push({ index: i, row: rowObj, range });
      }
    });
    return matches;
  }
}

// ========================= Typed Layer =========================
class SchemaValidator {
  static coerce(type: ColumnType | undefined, value: any): any {
    if (value == null || type == null || type === "any") return value;
    switch (type) {
      case "string":
        return value == null ? value : String(value);
      case "number": {
        if (value instanceof Date) return value.getTime();
        const n = Number(value);
        if (Number.isNaN(n)) throw new Error(`Cannot coerce value "${value}" to number`);
        return n;
      }
      case "boolean":
        if (typeof value === "boolean") return value;
        if (typeof value === "number") return value !== 0;
        if (typeof value === "string") return /^(true|1|yes|y)$/i.test(value.trim());
        return Boolean(value);
      case "date": {
        if (value instanceof Date) return value;
        // Excel might provide serial numbers or ISO strings
        if (typeof value === "number") {
          // Excel date serial number (assuming Windows 1900 system)
          const excelEpoch = new Date(Date.UTC(1899, 11, 30));
          const ms = value * 24 * 60 * 60 * 1000;
          return new Date(excelEpoch.getTime() + ms);
        }
        const d = new Date(value);
        if (isNaN(d.getTime())) throw new Error(`Cannot coerce value "${value}" to date`);
        return d;
      }
      default:
        return value;
    }
  }

  static defaultValue<T>(dv: DefaultValue<T> | undefined): T | undefined {
    if (dv === undefined) return undefined;
    return typeof dv === "function" ? (dv as () => T)() : dv;
  }
}

/** Maps typed keys to Excel headers & column constraints */
class ColumnMapping<T extends Record<string, any>> {
  readonly headers: string[];
  readonly keyByHeader: Map<string, keyof T> = new Map();
  readonly headerByKey: Map<keyof T, string> = new Map();
  readonly types: Map<keyof T, ColumnType | undefined> = new Map();
  readonly defaults: Map<keyof T, DefaultValue<any> | undefined> = new Map();
  readonly required: Set<keyof T> = new Set();
  readonly finals: Set<keyof T> = new Set();
  readonly calculated: Set<keyof T> = new Set();

  constructor(headers: string[], def: TableDefinition<T>) {
    this.headers = headers.slice();
    (Object.keys(def.columns) as (keyof T)[]).forEach((k) => {
      const header = (def.names?.[k] as string) ?? (k as string);
      this.headerByKey.set(k, header);
      this.types.set(k, def.columns[k].type);
      this.defaults.set(k, def.columns[k].default);
      if (def.columns[k].required) this.required.add(k);
      if ((def.columns[k] as any).final) this.finals.add(k);
      if ((def.columns[k] as any).calculated) this.calculated.add(k);
    });
    headers.forEach((h) => {
      const entry = (Object.keys(def.columns) as (keyof T)[]).find((k) => this.headerByKey.get(k) === h);
      if (entry) this.keyByHeader.set(h, entry);
    });
  }
}

export class TypedTable<T extends Record<string, any>> {
  readonly base: ExcelTable;
  readonly def: TableDefinition<T>;
  public readonly rows: TypedRowRepository<T>;

  constructor(base: ExcelTable, def: TableDefinition<T>) {
    this.base = base;
    this.def = def;
    this.rows = new TypedRowRepository<T>(base.native, def);
  }

  get native(): Excel.Table { return this.base.native; }
  async getSchema(): Promise<TableSchema> { return this.base.getSchema(); }
}

export class TypedRowRepository<T extends Record<string, any>> extends RowRepository {
  private readonly def: TableDefinition<T>;
  constructor(table: Excel.Table, def: TableDefinition<T>) { super(table); this.def = def; }

  private async mapping(): Promise<ColumnMapping<T>> {
    const headers = await this["headers"]();
    return new ColumnMapping<T>(headers, this.def);
  }

  /** Read and validate/coerce values to T */
  async getAll(options?: { coerce?: boolean }): Promise<T[]> {
    const coerce = options?.coerce !== false; // default true
    const headers = await this["headers"]();

    const body = (this as any)["table"].getDataBodyRange();
    body.load(["values"]);
    await ((this as any)["table"].context as Excel.RequestContext).sync();

    const mapping = await this.mapping();
    const rows: any[][] = body.values || [];

    const result: T[] = [];
    for (const r of rows) {
      const rec: any = {};
      headers.forEach((h, i) => {
        const key = mapping.keyByHeader.get(h);
        if (!key) return; // header not in definition — ignore
        const raw = r[i];
        rec[key as string] = coerce ? SchemaValidator.coerce(mapping.types.get(key), raw) : raw;
      });
      // required checks (ignore calculated)
      mapping.required.forEach((k) => {
        if (mapping.calculated.has(k)) return;
        const v = rec[k as string];
        if (isBlank(v)) throw new Error(`Required column missing/empty: ${String(k)}`);
      });
      result.push(rec as T);
    }
    return result;
  }

  /** Append typed row with defaults & coercion (skips calculated columns) */
  async add(obj: Partial<T>, options?: { fillDefaults?: boolean; coerce?: boolean }): Promise<void> {
    const { fillDefaults = true, coerce = true } = options ?? {};
    const headers = await this["headers"]();
    const mapping = await this.mapping();

    const rowObj: Record<string, any> = {};

    await excelLog("before Autoincrement");

    // Autoincrement
    const idKey: keyof T = 'id' as keyof T;
    if (isBlank((obj as any)[idKey])) {
      const maxId = await this.getMaxId();
      (obj as any)[idKey] = maxId + 1;
    }
    await excelLog("before headers.forEach maxId = " + (obj as any)[idKey]);

    headers.forEach((h) => {
      const key = mapping.keyByHeader.get(h);
      if (!key) { rowObj[h] = undefined; return; }
      if (mapping.calculated.has(key)) { rowObj[h] = undefined; return; } // never write calculated
      let v = (obj as any)[key];
      if (isBlank(v) && fillDefaults) {
        const dv = SchemaValidator.defaultValue(mapping.defaults.get(key));
        if (dv !== undefined) v = dv;
      }
      if (coerce) v = SchemaValidator.coerce(mapping.types.get(key), v);
      rowObj[h] = v;
    });
    // required validation (ignore calculated)
    mapping.required.forEach((k) => {
      if (mapping.calculated.has(k)) return;
      const header = mapping.headerByKey.get(k) as string;
      const v = rowObj[header];
      if (isBlank(v)) throw new Error(`Required value missing for ${String(k)} (${header})`);
    });

    excelLog("rowObj before add : " + JSON.stringify(rowObj));
    await super.add(rowObj);
  }  
  
  /** Bulk-insert: add many typed rows efficiently (single mapping/headers, chunked writes). */
  async addMany(objs: Partial<T>[], options?: { fillDefaults?: boolean; coerce?: boolean; chunkSize?: number }): Promise<number> {
    if (!objs || objs.length === 0) return 0;
    const { fillDefaults = true, coerce = true, chunkSize = 500 } = options ?? {};

    const headers = await this["headers"]();
    const mapping = await this.mapping();
    const table: Excel.Table = (this as any)["table"];

    const rowsToAdd: any[][] = [];
    
    let currentId = await this.getMaxId() + 1;

    for (const obj of objs) {
      // Autoincrement
      const idKey: keyof T = 'id' as keyof T;
      if (isBlank((obj as any)[idKey])) {
        (obj as any)[idKey] = currentId++;
      }

      const rowObj: Record<string, any> = {};
      headers.forEach((h) => {
        const key = mapping.keyByHeader.get(h);
        if (!key) { rowObj[h] = undefined; return; }
        if (mapping.calculated.has(key)) { rowObj[h] = undefined; return; } // skip calculated
        let v = (obj as any)[key];
        if (isBlank(v) && fillDefaults) {
          const dv = SchemaValidator.defaultValue(mapping.defaults.get(key));
          if (dv !== undefined) v = dv;
        }
        if (coerce) v = SchemaValidator.coerce(mapping.types.get(key), v);
        rowObj[h] = v;
      });
      // required validation (ignore calculated)
      mapping.required.forEach((k) => {
        if (mapping.calculated.has(k)) return;
        const header = mapping.headerByKey.get(k) as string;
        const v = rowObj[header];
        if (isBlank(v)) throw new Error(`Required value missing for ${String(k)} (${header})`);
      });
      rowsToAdd.push(objectToRow(headers, rowObj));
    }

    // Write in chunks to keep the command payload manageable
    for (let i = 0; i < rowsToAdd.length; i += chunkSize) {
      const chunk = rowsToAdd.slice(i, i + chunkSize);
      table.rows.add(null, chunk);
      await (table.context as Excel.RequestContext).sync();
    }
    return rowsToAdd.length;
  }

  /** Replace all rows with typed collection (preserves existing finals for matched rows) */
  async setAll(objs: Partial<T>[], options?: { fillDefaults?: boolean; coerce?: boolean }): Promise<void> {
    const { fillDefaults = true, coerce = true } = options ?? {};
    const headers = await this["headers"]();
    const mapping = await this.mapping();

    const table: Excel.Table = (this as any)["table"];
    const existingBody: any = (table as any).getDataBodyRangeOrNullObject ? (table as any).getDataBodyRangeOrNullObject() : table.getDataBodyRange();
    existingBody.load(["values", "rowCount", "isNullObject"]);
    await (table.context as Excel.RequestContext).sync();
    const existing: any[][] = Boolean(existingBody.isNullObject) ? [] : (existingBody.values || []);
    const existingCount = Boolean(existingBody.isNullObject) ? 0 : (existingBody.rowCount || existing.length);

    const rows = objs.map((obj, i) => {
      const rowObj: Record<string, any> = {};
      headers.forEach((h, colIndex) => {
        const key = mapping.keyByHeader.get(h);
        if (!key) { rowObj[h] = undefined; return; }
        if (mapping.calculated.has(key)) { rowObj[h] = undefined; return; }
        if (mapping.finals.has(key) && i < existingCount) {
          const prev = existing[i]?.[colIndex];
          if (!isBlank(prev)) { rowObj[h] = prev; return; }
        }
        let v = (obj as any)[key];
        if (isBlank(v) && fillDefaults) {
          const dv = SchemaValidator.defaultValue(mapping.defaults.get(key));
          if (dv !== undefined) v = dv;
        }
        if (coerce) v = SchemaValidator.coerce(mapping.types.get(key), v);
        rowObj[h] = v;
      });
      mapping.required.forEach((k) => {
        if (mapping.calculated.has(k)) return;
        const header = mapping.headerByKey.get(k) as string;
        const v = rowObj[header];
        if (isBlank(v)) throw new Error(`Required value missing for ${String(k)} (${header})`);
      });
      return rowObj;
    });

    await super.setAll(rows);
  }

  /** Find first row by typed key/value */
  async findFirstBy<K extends keyof T>(key: K, value: T[K], opts?: RowFindOptions): Promise<RowMatchTyped<T> | null> {
    const headers = await this["headers"]();
    const mapping = await this.mapping();
    const header = mapping.headerByKey.get(key);
    if (!header) throw new Error(`Key not mapped to a header: ${String(key)}`);

    const table: Excel.Table = (this as any)["table"];
    const colRange = table.columns.getItem(header).getDataBodyRange();
    colRange.load(["values", "rowCount"]);
    await (table.context as Excel.RequestContext).sync();
    if ((colRange as any).rowCount === 0) return null;

    const targetType = mapping.types.get(key);
    const target = SchemaValidator.coerce(targetType, value as any);

    const colValues: any[] = (colRange.values || []).map((r: any[]) => r[0]);
    const idx = colValues.findIndex((raw) => {
      const coerced = SchemaValidator.coerce(targetType, raw);
      return isEqual(coerced, target, opts);
    });
    if (idx === -1) return null;

    const body: any = (table as any).getDataBodyRangeOrNullObject ? (table as any).getDataBodyRangeOrNullObject() : table.getDataBodyRange();
    body.load(["values", "isNullObject"]);
    await (table.context as Excel.RequestContext).sync();

    const r = (body.values || [])[idx] as any[];
    const rec: any = {};
    headers.forEach((h, i) => {
      const k = mapping.keyByHeader.get(h);
      if (!k) return;
      const raw = r[i];
      rec[k as string] = SchemaValidator.coerce(mapping.types.get(k), raw);
    });

    const range = table.rows.getItemAt(idx).getRange();
    mapping.required.forEach((k) => {
      if (mapping.calculated.has(k)) return;
      const v = rec[k as string];
      if (isBlank(v)) throw new Error(`Required column missing/empty: ${String(k)}`);
    });

    return { index: idx, row: rec as T, range };
  }

  /** Find all rows by typed key/value */
  async findAllBy<K extends keyof T>(key: K, value: T[K], opts?: RowFindOptions): Promise<RowMatchTyped<T>[]> {
    const headers = await this["headers"]();
    const mapping = await this.mapping();
    const header = mapping.headerByKey.get(key);
    if (!header) throw new Error(`Key not mapped to a header: ${String(key)}`);

    const table: Excel.Table = (this as any)["table"];
    const colRange = table.columns.getItem(header).getDataBodyRange();
    colRange.load(["values", "rowCount"]);
    await (table.context as Excel.RequestContext).sync();

    const targetType = mapping.types.get(key);
    const target = SchemaValidator.coerce(targetType, value as any);

    const colValues: any[] = (colRange.values || []).map((r: any[]) => r[0]);

    const body = table.getDataBodyRange();
    body.load(["values"]);
    await (table.context as Excel.RequestContext).sync();

    const out: RowMatchTyped<T>[] = [];
    colValues.forEach((raw, idx) => {
      const coerced = SchemaValidator.coerce(targetType, raw);
      if (isEqual(coerced, target, opts)) {
        const r = (body.values || [])[idx] as any[];
        const rec: any = {};
        headers.forEach((h, i) => {
          const k = mapping.keyByHeader.get(h);
          if (!k) return;
          rec[k as string] = SchemaValidator.coerce(mapping.types.get(k), r[i]);
        });
        const range = table.rows.getItemAt(idx).getRange();
        out.push({ index: idx, row: rec as T, range });
      }
    });

    return out;
  }

  /** Update all rows that match key=value using a patch object. Respects final/calculated. */
  async updateBy<K extends keyof T>(key: K, value: T[K], patch: Partial<T>, options?: { fillDefaults?: boolean; coerce?: boolean }): Promise<number> {
    void options;
    const matches = await this.findAllBy(key, value);
    if (matches.length === 0) return 0;

    const headers = await this["headers"]();
    const mapping = await this.mapping();
    const table: Excel.Table = (this as any)["table"];

    for (const m of matches) {
      // Merge existing + patch honoring constraints
      const rowObj: Record<string, any> = {};
      headers.forEach((h) => {
        const k = mapping.keyByHeader.get(h);
        if (!k) { rowObj[h] = undefined; return; }
        if (mapping.calculated.has(k)) { rowObj[h] = undefined; return; }
        const incoming = (patch as any)[k];
        const current = (m.row as any)[k];
        if (mapping.finals.has(k)) {
          // Final cannot change if has a non-blank current value
          if (!isBlank(current) && !isBlank(incoming) && normalizeForCompare(current) !== normalizeForCompare(incoming)) {
            throw new Error(`Attempt to modify final column ${String(k)} on row ${m.index}`);
          }
          rowObj[h] = SchemaValidator.coerce(mapping.types.get(k), current);
          return;
        }
        let v = (incoming !== undefined ? incoming : current);
        v = SchemaValidator.coerce(mapping.types.get(k), v);
        rowObj[h] = v;
      });
      // Required check (ignore calculated)
      mapping.required.forEach((k) => {
        if (mapping.calculated.has(k)) return;
        const header = mapping.headerByKey.get(k) as string;
        const v = rowObj[header];
        if (isBlank(v)) throw new Error(`Required value missing for ${String(k)} (${header}) on row ${m.index}`);
      });
      const rowRange = table.rows.getItemAt(m.index).getRange();
      rowRange.values = [objectToRow(headers, rowObj)];
      await (table.context as Excel.RequestContext).sync();
    }

    return matches.length;
  }

  /** Delete all rows where key=value. */
  async deleteBy<K extends keyof T>(key: K, value: T[K], opts?: RowFindOptions): Promise<number> {
    // opts is used in findAllBy

    const matches = await this.findAllBy(key, value, opts);
    if (matches.length === 0) return 0;
    const table: Excel.Table = (this as any)["table"];
    // Delete from bottom to top to keep indices valid
    const indices = matches.map((m) => m.index).sort((a, b) => b - a);
    for (const idx of indices) {
      table.rows.getItemAt(idx).delete();
      await (table.context as Excel.RequestContext).sync();
    }
    return matches.length;
  }

  async getMaxId(): Promise<number> {
    const mapping = await this.mapping();
    const idKey: keyof T = 'id' as keyof T; // Предполагаем ключ 'id'; можно сделать параметром если нужно
  
    if (!mapping.headerByKey.has(idKey)) {
      throw new Error(`No column mapped for key "${String(idKey)}"`);
    }
  
    const type = mapping.types.get(idKey);
    if (type !== 'number') {
      throw new Error(`ID column "${String(idKey)}" must be of type "number"`);
    }
  
    const header = mapping.headerByKey.get(idKey)!;
    const table: Excel.Table = (this as any)["table"]; // this.table из RowRepository
  
    const colRange = table.columns.getItem(header).getDataBodyRange();
    colRange.load(["values", "rowCount"]);
    await table.context.sync();
  
    if (colRange.rowCount === 0) {
      return 0; // Таблица пустая, начинаем с 1
    }
  
    const values: number[] = colRange.values
      .flat()
      .map(v => Number(v))
      .filter(v => !isNaN(v) && Number.isFinite(v));
  
    const max = values.length > 0 ? Math.max(...values) : 0;
    return max;
  }  
}

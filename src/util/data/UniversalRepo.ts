/// <reference types="office-js" />
/*
  Excel Workbook ORM (metadata + lightweight accessors) — now with Generics & Typed Rows
  --------------------------------------------------------------------------------------
  What's new:
  - TableRepository.getAs<T>(name, definition) → returns a TypedTable<T>
  - Column typing & validation: string | number | boolean | date | any
  - Object <-> row mapping with optional name mapping, default values, coercion & required checks

  Quick start:

  type User = { Id: number; Name: string; Email: string; IsActive: boolean; CreatedAt: Date };

  const userDef: TableDefinition<User> = {
    columns: {
      Id:       { type: "number",  required: true },
      Name:     { type: "string",  required: true },
      Email:    { type: "string" },
      IsActive: { type: "boolean", default: true },
      CreatedAt:{ type: "date",    default: () => new Date() },
    },
    // Optional: map keys to Excel header names if they differ (default = same name)
    names: { /* Email: "E-mail" * / },
    // Optional: control column order when writing (defaults to header order)
    order: ["Id", "Name", "Email", "IsActive", "CreatedAt"],
  };

  await Excel.run(async (ctx) => {
    const orm = new WorkbookORM(ctx.workbook);

    const users = await orm.tables.getAs<User>("Users", userDef);

    // Read typed records
    const list: User[] = await users.rows.getAll();

    // Add typed row with validation/coercion/defaults
    await users.rows.add({ Id: 101, Name: "Alice", Email: "a@ex.com" });
  });
*/

// ========================= Shared Types =========================

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

export type ColumnType = "string" | "number" | "boolean" | "date" | "any";
export type DefaultValue<T> = T | (() => T);

export type TableDefinition<T extends Record<string, any>> = {
  columns: {
    [K in keyof T]-?: {
      type?: ColumnType;
      required?: boolean;
      default?: DefaultValue<T[K]>;
      // добавили метаданные из шапки
      calculated?: boolean;
      referenceTo?: string;
      final?: boolean;
    };
  };
  names?: Partial<Record<keyof T, string>>;
  order?: (keyof T)[];
};
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

  constructor(workbook: Excel.Workbook) {
    this.workbook = workbook;
  }

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
      const columns: ColumnSchema[] = t.columns.items.map((c) => ({
        name: c.name,
        index: c.index,
      }));
      results.push({
        id: t.id,
        name: t.name,
        worksheet: t.worksheet.name,
        columns,
      });
    }
    return results.sort((a, b) => a.name.localeCompare(b.name));
  }

  /** Get an untyped wrapped table by name (case-sensitive Excel name) */
  async get(name: string): Promise<ExcelTable> {
    const cached = this.cacheByName.get(name);
    if (cached) return cached;

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
    return {
      id: t.id,
      name: t.name,
      worksheet: t.worksheet.name,
      columns,
    };
  }

  /** Native reference */
  get native(): Excel.Table {
    return this.table;
  }
}

// ========================= Column Access (Untyped) =========================
export class ColumnRepository {
  private table: Excel.Table;
  private byName: Map<string, ExcelColumn> = new Map();

  constructor(table: Excel.Table) {
    this.table = table;
  }

  async list(): Promise<ExcelColumn[]> {
    const ctx = this.table.context as Excel.RequestContext;
    this.table.columns.load(["items/name", "items/index"]);
    await ctx.sync();
    return this.table.columns.items
      .sort((a, b) => a.index - b.index)
      .map((col) => new ExcelColumn(this.table, col.name, col.index));
  }

  async get(name: string): Promise<ExcelColumn> {
    const cached = this.byName.get(name);
    if (cached) return cached;

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

  constructor(table: Excel.Table, name: string, index: number) {
    this.table = table;
    this.name = name;
    this.index = index;
  }

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

    if (rng.rowCount !== values.length) {
      throw new Error(
        `Length mismatch for column "${this.name}": table has ${rng.rowCount} rows, received ${values.length}`
      );
    }
    rng.values = values.map((v) => [v]);
    await (this.table.context as Excel.RequestContext).sync();
  }
}

// ========================= Row Helpers (Untyped) =========================
export class RowRepository {
  private table: Excel.Table;

  constructor(table: Excel.Table) {
    this.table = table;
  }

  protected async headers(): Promise<string[]> {
    const headerRange = this.table.getHeaderRowRange();
    headerRange.load(["values"]);
    await (this.table.context as Excel.RequestContext).sync();
    const headers = (headerRange.values?.[0] || []) as string[];
    return headers.map((h) => String(h));
  }

  async getAll(): Promise<Record<string, any>[]> {
    const headers = await this.headers();
    const body = this.table.getDataBodyRange();
    body.load(["values"]);
    await (this.table.context as Excel.RequestContext).sync();
    const values: any[][] = body.values || [];
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

    if (body.rowCount > 0) {
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
    if (body.rowCount > 0) {
      body.clear(Excel.ClearApplyTo.contents);
      await (this.table.context as Excel.RequestContext).sync();
    }
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

/**
 * Maps typed keys to actual Excel headers and stores column type/defaults.
 */
class ColumnMapping<T extends Record<string, any>> {
  readonly headers: string[];
  readonly keyByHeader: Map<string, keyof T> = new Map();
  readonly headerByKey: Map<keyof T, string> = new Map();
  readonly types: Map<keyof T, ColumnType | undefined> = new Map();
  readonly defaults: Map<keyof T, DefaultValue<any> | undefined> = new Map();
  readonly required: Set<keyof T> = new Set();

  constructor(headers: string[], def: TableDefinition<T>) {
    this.headers = headers.slice();

    // Resolve header names for each key
    (Object.keys(def.columns) as (keyof T)[]).forEach((k) => {
      const header = (def.names?.[k] as string) ?? (k as string);
      this.headerByKey.set(k, header);
      this.types.set(k, def.columns[k].type);
      this.defaults.set(k, def.columns[k].default);
      if (def.columns[k].required) this.required.add(k);
    });

    // Build reverse map only for headers that exist
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

  constructor(table: Excel.Table, def: TableDefinition<T>) {
    super(table);
    this.def = def;
  }

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
      // required checks
      mapping.required.forEach((k) => {
        const v = rec[k as string];
        if (v === undefined || v === null || v === "") {
          throw new Error(`Required column missing/empty: ${String(k)}`);
        }
      });
      result.push(rec as T);
    }
    return result;
  }

  /** Append typed row with defaults & coercion */
  async add(obj: Partial<T>, options?: { fillDefaults?: boolean; coerce?: boolean }): Promise<void> {
    const { fillDefaults = true, coerce = true } = options ?? {};
    const headers = await this["headers"]();
    const mapping = await this.mapping();

    const rowObj: Record<string, any> = {};

    // Prepare by reading def over headers order
    headers.forEach((h) => {
      const key = mapping.keyByHeader.get(h);
      if (!key) {
        // header not in def — write raw if provided by name
        rowObj[h] = undefined;
        return;
      }
      let v = (obj as any)[key];
      if ((v === undefined || v === null || v === "") && fillDefaults) {
        const dv = SchemaValidator.defaultValue(mapping.defaults.get(key));
        if (dv !== undefined) v = dv;
      }
      if (coerce) v = SchemaValidator.coerce(mapping.types.get(key), v);
      rowObj[h] = v;
    });

    // Required validation
    mapping.required.forEach((k) => {
      const header = mapping.headerByKey.get(k) as string;
      const v = rowObj[header];
      if (v === undefined || v === null || v === "") {
        throw new Error(`Required value missing for ${String(k)} (${header})`);
      }
    });

    await super.add(rowObj);
  }

  /** Replace all rows with typed collection */
  async setAll(objs: Partial<T>[], options?: { fillDefaults?: boolean; coerce?: boolean }): Promise<void> {
    const { fillDefaults = true, coerce = true } = options ?? {};
    const headers = await this["headers"]();
    const mapping = await this.mapping();

    const rows = objs.map((obj) => {
      const rowObj: Record<string, any> = {};
      headers.forEach((h) => {
        const key = mapping.keyByHeader.get(h);
        if (!key) { rowObj[h] = undefined; return; }
        let v = (obj as any)[key];
        if ((v === undefined || v === null || v === "") && fillDefaults) {
          const dv = SchemaValidator.defaultValue(mapping.defaults.get(key));
          if (dv !== undefined) v = dv;
        }
        if (coerce) v = SchemaValidator.coerce(mapping.types.get(key), v);
        rowObj[h] = v;
      });
      mapping.required.forEach((k) => {
        const header = mapping.headerByKey.get(k) as string;
        const v = rowObj[header];
        if (v === undefined || v === null || v === "") {
          throw new Error(`Required value missing for ${String(k)} (${header})`);
        }
      });
      return rowObj;
    });

    await super.setAll(rows);
  }
}

/// <reference types="office-js" />
/*
  WorkbookSchemaGenerator (supports metadata rows above table header)
  ------------------------------------------------------------------
  Reads Excel tables and optional metadata block placed ABOVE the table header, aligned to columns.
  The label for each metadata row is in the column immediately to the left of the table (e.g., A column).

  Supported metadata labels (case-insensitive in the left label column):
    Type, Required, Calculated, ReferenceTo, Final, DefaultValue

  Mapping is done by label text, not by hard-coded order.
*/

export type ColumnType = "string" | "number" | "boolean" | "date" | "any";
export type DefaultValue<T> = T | (() => T);
export type TableDefinition<T extends Record<string, any>> = {
  columns: {
    [K in keyof T]-?: {
      type?: ColumnType;
      required?: boolean;
      default?: DefaultValue<T[K]>;
      calculated?: boolean;
      // referenceTo?: string;
      final?: boolean;
    }
  };
  names?: Partial<Record<keyof T, string>>;
  order?: (keyof T)[];
};

export type GeneratedTable = {
  tableName: string;
  typeName: string;
  constName: string;
  keys: string[];
  headers: string[];
  namesMap?: Record<string, string>;
  def: TableDefinition<Record<string, any>>;
  code: string;
};

export type GenerateOptions = {
  sampleRows?: number;
  emitHeader?: boolean;
  inlineTableDefinition?: boolean;
  importPath?: string;
};

const META_KEYS = ["type", "required", "calculated", "referenceto", "final", "defaultvalue"] as const;

type MetaMap = Partial<Record<(typeof META_KEYS)[number], any[]>>;

export class WorkbookSchemaGenerator {
  constructor(private workbook: Excel.Workbook) {}

  async scan(sampleRows = 50): Promise<Array<{ table: Excel.Table; headers: string[]; samples: any[][]; meta: MetaMap }>> {
    const ctx = this.workbook.context as Excel.RequestContext;
    const tables = this.workbook.tables;
    tables.load(["items"]);
    await ctx.sync();

    for (const t of tables.items) {
      t.load(["name", "id", "worksheet/name"]);
    }
    await ctx.sync();

    const result: Array<{ table: Excel.Table; headers: string[]; samples: any[][]; meta: MetaMap }> = [];

    for (const t of tables.items) {
      const ws = t.worksheet;
      const header = t.getHeaderRowRange();
      header.load(["values", "rowIndex", "columnIndex", "columnCount"]);
      const body = t.getDataBodyRange();
      body.load(["values", "rowCount"]);
      await ctx.sync();

      const headers = (header.values?.[0] || []).map((v) => String(v));
      const samples = body.values && body.values.length > 0 ? body.values.slice(0, sampleRows) : [];
      const meta = await this.readMetadataBlock(ws, header);

      result.push({ table: t, headers, samples, meta });
    }
    return result;
  }

  private async readMetadataBlock(ws: Excel.Worksheet, header: Excel.Range): Promise<MetaMap> {
    const ctx = ws.context as Excel.RequestContext;
    header.load(["rowIndex", "columnIndex", "columnCount"]);
    await ctx.sync();

    const labelCol = Math.max(0, header.columnIndex - 1);
    const startRow = Math.max(0, header.rowIndex - 50);
    const rowCount = header.rowIndex - startRow;
    if (rowCount <= 0) return {};

    const metaRange = ws.getRangeByIndexes(startRow, labelCol, rowCount, header.columnCount + 1);
    metaRange.load(["values"]);
    await ctx.sync();

    const meta: MetaMap = {};
    const values: any[][] = metaRange.values || [];

    // bottom-up: collect known labels
    for (let r = values.length - 1; r >= 0; r--) {
      const label = String(values[r]?.[0] ?? "").trim().toLowerCase();
      if (!label || !(META_KEYS as readonly string[]).includes(label)) continue;
      meta[label as (typeof META_KEYS)[number]] = values[r].slice(1);
    }
    return meta;
  }

  async generateTypeScript(options?: GenerateOptions): Promise<{ code: string; runtime: Record<string, TableDefinition<Record<string, any>>>; tables: GeneratedTable[] }> {
    const opt: Required<GenerateOptions> = {
      sampleRows: options?.sampleRows ?? 50,
      emitHeader: options?.emitHeader ?? true,
      inlineTableDefinition: options?.inlineTableDefinition ?? false,
      importPath: options?.importPath ?? "./excel-orm",
    } as Required<GenerateOptions>;

    const scanned = await this.scan(opt.sampleRows);
    const generated: GeneratedTable[] = [];
    const runtime: Record<string, TableDefinition<Record<string, any>>> = {};

    for (const s of scanned) {
      const typeName = toPascalCase(s.table.name);
      const constName = toCamelCase(typeName) + "Def";
      const { keys, headers, namesMap, columns, order } = this.inferForTable(s.headers, s.samples, s.meta);

      const def: TableDefinition<Record<string, any>> = { columns, ...(namesMap && Object.keys(namesMap).length ? { names: namesMap as any } : {}), order: order as any };
      const code = this.emitTableCode({ typeName, constName, keys, headers, namesMap, def });

      generated.push({ tableName: s.table.name, typeName, constName, keys, headers, namesMap, def, code });
      runtime[s.table.name] = def;
    }

    const header = opt.emitHeader ? `// Auto-generated by WorkbookSchemaGenerator on ${new Date().toISOString()}\n// Tables: ${generated.map((g) => g.tableName).join(", ")}\n` : "";
    const importOrInline = opt.inlineTableDefinition ? inlineTableDefinitionBlock() : `import { TableDefinition } from "${opt.importPath}";\n`;
    const code = `${header}${importOrInline}\n${generated.map((g) => g.code).join("\n\n")}\n`;

    return { code, runtime, tables: generated };
  }

  inferForTable(
    headers: string[],
    samples: any[][],
    meta: MetaMap
  ): {
    keys: string[];
    headers: string[];
    namesMap?: Record<string, string>;
    columns: Record<string, { type?: ColumnType; required?: boolean; default?: any; calculated?: boolean; referenceTo?: string; final?: boolean }>;
    order: string[];
  } {
    const keys = headers.map((h) => sanitizeIdentifier(h));

    const namesMap: Record<string, string> = {};
    keys.forEach((k, i) => { const h = headers[i]; if (k !== h) namesMap[k] = h; });

    const typeRow = meta["type"];
    const colTypes: ColumnType[] = headers.map((h, i) => {
      const m = typeRow?.[i];
      if (m != null && String(m).trim() !== "") return normalizeType(String(m));
      const hg = guessTypeFromHeader(h);
      if (hg !== "any") return hg;
      const vals = samples.map((r) => r[i]).filter((v) => v !== undefined && v !== null && v !== "");
      return guessTypeFromValues(vals);
    });

    const reqRow = meta["required"];     // boolean per col
    const calcRow = meta["calculated"];  // boolean per col
    const finRow = meta["final"];        // boolean per col
    const refRow = meta["referenceto"];  // string per col
    const defRow = meta["defaultvalue"]; // any per col

    const columns: Record<string, { type?: ColumnType; required?: boolean; default?: any; calculated?: boolean; referenceTo?: string; final?: boolean }> = {};

    keys.forEach((k, i) => {
      const t = colTypes[i];
      const col: any = { type: t };
      if (truthy(reqRow?.[i])) col.required = true;
      if (truthy(calcRow?.[i])) col.calculated = true;
      if (truthy(finRow?.[i])) col.final = true;
      const ref = asNonEmptyString(refRow?.[i]); if (ref) col.referenceTo = ref;
      const dv = defValue(defRow?.[i], t); if (dv !== undefined) col.default = dv;
      columns[k] = col;
    });

    return { keys, headers, namesMap: Object.keys(namesMap).length ? namesMap : undefined, columns, order: keys.slice() };
  }

  private emitTableCode(
    g: { typeName: string; constName: string; keys: string[]; headers: string[]; namesMap?: Record<string, string>; def: TableDefinition<Record<string, any>> }
  ): string {
    const iface = `export interface ${g.typeName} {\n${g.keys.map((k) => `  ${k}: ${toTsFieldType((g.def.columns as any)[k]?.type)};`).join("\n")}\n}`;

    const columnsBlock = g.keys.map((k) => {
      const c = (g.def.columns as any)[k] || {};
      const parts: string[] = [];
      if (c.type) parts.push(`type: "${c.type}"`);
      if (c.required) parts.push(`required: true`);
      if (c.calculated) parts.push(`calculated: true`);
      if (c.final) parts.push(`final: true`);
      if (c.referenceTo) parts.push(`referenceTo: ${JSON.stringify(c.referenceTo)}`);
      if (typeof c.default === "function") parts.push(`default: () => new Date()`);
      else if (c.default !== undefined) parts.push(`default: ${formatDefaultForCode(c.default, c.type)}`);
      return `    ${k}: { ${parts.join(", ")} }`;
    }).join(",\n");

    const namesBlock = g.namesMap ? `,\n  names: {\n${Object.entries(g.namesMap).map(([k, h]) => `    ${k}: ${JSON.stringify(h)}`).join(",\n")}\n  }` : "";
    const orderBlock = `,\n  order: [${g.keys.map((k) => JSON.stringify(k)).join(", ")}]`;

    const def = `export const ${g.constName}: TableDefinition<${g.typeName}> = {\n  columns: {\n${columnsBlock}\n  }${namesBlock}${orderBlock}\n};`;
    return `${iface}\n\n${def}`;
  }
}

// ---------------- helpers ----------------
function toPascalCase(s: string): string {
  const parts = s.replace(/[^A-Za-z0-9]+/g, " ").trim().split(/\s+/);
  let out = parts.map((p) => p.charAt(0).toUpperCase() + p.slice(1)).join("");
  if (!/^[A-Za-z_]/.test(out)) out = "T" + out; return out;
}
function toCamelCase(s: string): string { const p = toPascalCase(s); return p.charAt(0).toLowerCase() + p.slice(1); }
function sanitizeIdentifier(s: string): string { const cc = toCamelCase(s); return cc.replace(/[^A-Za-z0-9_]/g, "_"); }

function toTsFieldType(t?: ColumnType): string {
  switch (t) { case "number": return "number"; case "boolean": return "boolean"; case "date": return "Date"; case "string": return "string"; default: return "any"; }
}
function normalizeType(s: string): ColumnType {
  const v = s.trim().toLowerCase();
  if (["number", "numeric", "int", "integer", "float", "double"].includes(v)) return "number";
  if (["bool", "boolean", "true/false"].includes(v)) return "boolean";
  if (["date", "datetime", "timestamp"].includes(v)) return "date";
  if (["string", "text", "varchar"].includes(v)) return "string";
  return "any";
}
function guessTypeFromHeader(h: string): ColumnType {
  const s = h.toLowerCase();
  if (/^(id|#id|record ?id)$/.test(s) || /id$/.test(s)) return "number";
  if (/^(is|has)[ _-]?/.test(s) || /(active|enabled|disabled|archived)$/.test(s)) return "boolean";
  if (/(date|created|updated|time|at)$/.test(s)) return "date";
  if (/(email|phone|name|title|desc|address|city|country|zip)/.test(s)) return "string";
  return "any";
}
function guessTypeFromValues(values: any[]): ColumnType {
  for (const v of values) {
    if (v instanceof Date) return "date";
    if (typeof v === "boolean") return "boolean";
    if (typeof v === "number") return "number";
    if (typeof v === "string") {
      const s = v.trim(); if (s === "") continue;
      if (/^(true|false|yes|no|y|n|0|1)$/i.test(s)) return "boolean";
      const n = Number(s); if (!Number.isNaN(n) && Number.isFinite(n)) return "number";
      const d = new Date(s); if (!Number.isNaN(d.getTime())) return "date";
    }
  }
  return "string";
}
function truthy(v: any): boolean { if (v == null) return false; if (typeof v === "boolean") return v; if (typeof v === "number") return v !== 0; const s = String(v).trim().toLowerCase(); return s === "true" || s === "1" || s === "yes" || s === "y"; }
function asNonEmptyString(v: any): string | undefined { if (v == null) return undefined; const s = String(v).trim(); return s ? s : undefined; }
function defValue(v: any, t: ColumnType): any | undefined {
  if (v == null || v === "") return undefined;
  switch (t) {
    case "number": { const n = typeof v === "number" ? v : Number(v); return Number.isFinite(n) ? n : undefined; }
    case "boolean": return truthy(v);
    case "date": {
      if (v instanceof Date) return () => new Date(v.getTime());
      if (typeof v === "number") { const excelEpoch = new Date(Date.UTC(1899, 11, 30)); const ms = v * 86400000; const d = new Date(excelEpoch.getTime() + ms); return () => new Date(d.getTime()); }
      const d = new Date(v); return Number.isNaN(d.getTime()) ? undefined : (() => new Date(d.getTime()));
    }
    default: return String(v);
  }
}
function formatDefaultForCode(v: any, t?: ColumnType): string {
  if (v === undefined) return "undefined";
  if (t === "date" && typeof v === "string") { const d = new Date(v); if (!Number.isNaN(d.getTime())) return `() => new Date(${JSON.stringify(d.toISOString())})`; }
  if (typeof v === "string") return JSON.stringify(v);
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  return JSON.stringify(v);
}
function inlineTableDefinitionBlock(): string {
  return `export type ColumnType = "string" | "number" | "boolean" | "date" | "any";\nexport type DefaultValue<T> = T | (() => T);\nexport type TableDefinition<T extends Record<string, any>> = {\n  columns: { [K in keyof T]-?: { type?: ColumnType; required?: boolean; default?: DefaultValue<T[K]>; calculated?: boolean; referenceTo?: string; final?: boolean } };\n  names?: Partial<Record<keyof T, string>>;\n  order?: (keyof T)[];\n};\n`; 
}

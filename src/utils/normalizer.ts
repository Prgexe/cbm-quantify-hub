import * as XLSX from "xlsx";

export const BASE_FIELDS = [
  "QTD",
  "AREA",
  "UNIDADE",
  "POSTO_GRAD",
  "QUADRO",
  "NOME_COMPLETO",
  "RG",
] as const;

export type CanonicalBaseField = (typeof BASE_FIELDS)[number];

export const BASE_FIELD_LABELS: Record<CanonicalBaseField, string> = {
  QTD: "QTD",
  AREA: "ÁREA",
  UNIDADE: "UNIDADE",
  POSTO_GRAD: "POSTO/GRAD",
  QUADRO: "QUADRO",
  NOME_COMPLETO: "NOME COMPLETO",
  RG: "RG",
};

export interface ParsedImportRecord {
  QTD: string;
  AREA: string;
  UNIDADE: string;
  POSTO_GRAD: string;
  QUADRO: string;
  NOME_COMPLETO: string;
  RG: string;
  materiais: Record<string, string>;
}

export interface ParsedWorkbookResult {
  records: ParsedImportRecord[];
  warnings: string[];
  materials: string[];
  previewColumns: string[];
  previewRows: Array<Record<string, string>>;
}

type RowRecord = { materiais: Record<string, string> };

interface HeaderMatch {
  kind: "base" | "material";
  columnIndex: number;
  rawHeader: string;
  canonical: string;
  field?: CanonicalBaseField;
  score: number;
}

interface HeaderCandidate {
  score: number;
  baseCount: number;
  requiredCount: number;
  materialCount: number;
  headerRowIndex: number;
  dataStartRowIndex: number;
  matches: HeaderMatch[];
}

const REQUIRED_BASE_FIELDS: CanonicalBaseField[] = ["AREA", "UNIDADE", "POSTO_GRAD", "NOME_COMPLETO", "RG"];

const IGNORED_HEADER_LABELS = new Set([
  "",
  "TAMANHO",
  "TAMANHOS",
  "MATERIAL",
  "MATERIAIS",
  "OBS",
  "OBSERVACAO",
  "OBSERVACOES",
  "ASSINATURA",
]);

const HEADER_ALIASES: Record<CanonicalBaseField, string[]> = {
  QTD: ["QTD", "QTD.", "QUANTIDADE"],
  AREA: ["AREA", "ÁREA", "AREA FINAL", "ÁREA FINAL"],
  UNIDADE: ["UNIDADE", "UNIDADE FINAL", "ADE", "ADE FINAL", "ADE (FINAL)", "UNIDADE / ADE (FINAL)"],
  POSTO_GRAD: ["POSTO/GRAD", "POSTO / GRAD", "POSTO", "POSTO/GRADUAÇÃO", "POSTO/GRADUACAO", "GRADUAÇÃO", "GRADUACAO", "GRAD"],
  QUADRO: ["QUADRO", "QUAD", "QD"],
  NOME_COMPLETO: ["NOME COMPLETO", "NOME", "NOME DO MILITAR"],
  RG: ["RG", "REGISTRO", "N RG", "Nº RG", "NUMERO RG", "NÚMERO RG"],
};

const MATERIAL_ALIASES: Record<string, string[]> = {
  "CAMISETA GV": ["CAMISETA GV", "CAMISETA DE GV"],
  "CAMISA UV": ["CAMISA UV", "CAMISA U.V", "CAMISA U.V PARA GV", "CAMISA UV PARA GV"],
  "SHORT JOHN": ["SHORT JOHN"],
  "SHORT GMAR": ["SHORT GMAR", "SHORT G MAR", "SHORT"],
  "SUNGA GMAR": ["SUNGA GMAR", "SUNGA G MAR", "SUNGA"],
  BUSTIE: ["BUSTIE", "BUSTIÉ", "BUSTIE "],
  "SUNGA FEMININA": ["SUNGA FEMININA", "SUNGA FEM", "SUNGA FEM.", "SUNGA FEMIN."],
};

const POSTO_PATTERNS: Array<{ canonical: string; patterns: string[] }> = [
  { canonical: "CEL", patterns: ["CORONEL", " CEL ", "CEL"] },
  { canonical: "TEN CEL", patterns: ["TENENTE CORONEL", "TEN CEL", "TEN. CEL", "TENCEL", "TC"] },
  { canonical: "MAJ", patterns: ["MAJOR", "MAJ"] },
  { canonical: "CAP", patterns: ["CAPITAO", "CAPITÃO", "CAP"] },
  { canonical: "1º TEN", patterns: ["1 TENENTE", "1 TEN", "1ºTEN", "1º TEN", "PRIMEIRO TENENTE"] },
  { canonical: "2º TEN", patterns: ["2 TENENTE", "2 TEN", "2ºTEN", "2º TEN", "SEGUNDO TENENTE"] },
  { canonical: "TEN", patterns: ["TENENTE", "TEN"] },
  { canonical: "SUBTENENTE", patterns: ["SUBTENENTE", "SUB TEN", "SUBTEN", "ST"] },
  { canonical: "1º SGT", patterns: ["1 SGT", "1ºSGT", "1º SGT", "PRIMEIRO SARGENTO"] },
  { canonical: "2º SGT", patterns: ["2 SGT", "2ºSGT", "2º SGT", "SEGUNDO SARGENTO"] },
  { canonical: "3º SGT", patterns: ["3 SGT", "3ºSGT", "3º SGT", "TERCEIRO SARGENTO"] },
  { canonical: "CB", patterns: ["CABO", "CB"] },
  { canonical: "SD", patterns: ["SOLDADO", "SD"] },
];

const POSTO_ORDER: Record<string, number> = {
  CEL: 1,
  "TEN CEL": 2,
  MAJ: 3,
  CAP: 4,
  "1º TEN": 5,
  "2º TEN": 6,
  TEN: 7,
  SUBTENENTE: 8,
  "1º SGT": 9,
  "2º SGT": 10,
  "3º SGT": 11,
  CB: 12,
  SD: 13,
};

const SIZE_MAP: Record<string, string> = {
  PP: "PP",
  P: "P",
  M: "M",
  G: "G",
  GG: "GG",
  XG: "XG",
  XGG: "XG",
  XXG: "XG",
  EXG: "XG",
  EG: "XG",
  X: "X",
  NA: "X",
  NAOAPLICA: "X",
  NÃOAPLICA: "X",
  NAPLICA: "X",
  "--": "--",
  "-": "--",
};

const SIZE_ORDER = ["PP", "P", "M", "G", "GG", "XG", "X", "--"];

const SHEETS_TO_IGNORE = ["resumo geral", "contagem", "acrescentar1"];

function stripAccents(value: string) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function collapseSpaces(value: string) {
  return value.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function stringifyCell(value: unknown) {
  if (value == null) return "";
  if (typeof value === "number") return Number.isInteger(value) ? String(value) : String(value).replace(/\.0+$/, "");
  return String(value);
}

function normalizeLookup(value: unknown) {
  return stripAccents(collapseSpaces(stringifyCell(value)).toUpperCase())
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeDisplayText(value: unknown) {
  return collapseSpaces(stringifyCell(value)).toUpperCase();
}

function levenshtein(a: string, b: string) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;

  const matrix = Array.from({ length: b.length + 1 }, () => new Array<number>(a.length + 1).fill(0));
  for (let i = 0; i <= b.length; i += 1) matrix[i][0] = i;
  for (let j = 0; j <= a.length; j += 1) matrix[0][j] = j;

  for (let i = 1; i <= b.length; i += 1) {
    for (let j = 1; j <= a.length; j += 1) {
      const cost = a[j - 1] === b[i - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost,
      );
    }
  }

  return matrix[b.length][a.length];
}

function similarityScore(input: string, alias: string) {
  if (!input || !alias) return 0;
  if (input === alias) return 1;
  if (input.includes(alias) || alias.includes(input)) return 0.94;

  const inputTokens = new Set(input.split(" "));
  const aliasTokens = new Set(alias.split(" "));
  const overlap = [...inputTokens].filter((token) => aliasTokens.has(token)).length;
  const tokenScore = overlap / Math.max(inputTokens.size, aliasTokens.size, 1);
  const distance = levenshtein(input, alias);
  const editScore = 1 - distance / Math.max(input.length, alias.length, 1);

  return Math.max(tokenScore, editScore * 0.9);
}

function matchBaseField(header: string): { field: CanonicalBaseField; score: number } | null {
  const lookup = normalizeLookup(header);
  let best: { field: CanonicalBaseField; score: number } | null = null;

  for (const field of BASE_FIELDS) {
    for (const alias of HEADER_ALIASES[field]) {
      const score = similarityScore(lookup, normalizeLookup(alias));
      if (!best || score > best.score) best = { field, score };
    }
  }

  if (!best) return null;
  if (best.score >= 0.74) return best;
  if (lookup.includes("POSTO") || lookup.includes("GRAD")) return { field: "POSTO_GRAD", score: 0.8 };
  if (lookup.includes("UNIDADE") || lookup.includes("ADE")) return { field: "UNIDADE", score: 0.8 };
  if (lookup.includes("NOME")) return { field: "NOME_COMPLETO", score: 0.8 };
  return null;
}

function canonicalizeMaterialHeader(header: string): string | null {
  const display = normalizeDisplayText(header);
  const lookup = normalizeLookup(header);
  if (!lookup || IGNORED_HEADER_LABELS.has(lookup)) return null;
  if (matchBaseField(header)) return null;

  let best: { canonical: string; score: number } | null = null;
  for (const [canonical, aliases] of Object.entries(MATERIAL_ALIASES)) {
    for (const alias of aliases) {
      const score = similarityScore(lookup, normalizeLookup(alias));
      if (!best || score > best.score) best = { canonical, score };
    }
  }

  if (best && best.score >= 0.74) return best.canonical;
  if (/^[0-9]+$/.test(lookup)) return null;
  return display;
}

function mergeHeaderRows(top: unknown[] = [], bottom: unknown[] = []) {
  const length = Math.max(top.length, bottom.length);
  return Array.from({ length }, (_, index) => {
    const first = normalizeDisplayText(top[index]);
    const second = normalizeDisplayText(bottom[index]);

    if (!first) return second;
    if (!second) return first;
    if (IGNORED_HEADER_LABELS.has(normalizeLookup(first))) return second;
    if (IGNORED_HEADER_LABELS.has(normalizeLookup(second))) return first;
    if (normalizeLookup(first) === normalizeLookup(second)) return first;
    return `${first} ${second}`.trim();
  });
}

function evaluateHeaderCandidate(row: unknown[], rowIndex: number, dataStartRowIndex: number): HeaderCandidate | null {
  const matchesByBase = new Map<CanonicalBaseField, HeaderMatch>();
  const materialMatches = new Map<string, HeaderMatch>();

  row.forEach((cell, columnIndex) => {
    const rawHeader = normalizeDisplayText(cell);
    if (!rawHeader) return;

    const baseMatch = matchBaseField(rawHeader);
    if (baseMatch) {
      const match: HeaderMatch = {
        kind: "base",
        columnIndex,
        rawHeader,
        canonical: BASE_FIELD_LABELS[baseMatch.field],
        field: baseMatch.field,
        score: baseMatch.score,
      };
      const current = matchesByBase.get(baseMatch.field);
      if (!current || current.score < match.score) matchesByBase.set(baseMatch.field, match);
      return;
    }

    const material = canonicalizeMaterialHeader(rawHeader);
    if (!material) return;

    const match: HeaderMatch = {
      kind: "material",
      columnIndex,
      rawHeader,
      canonical: material,
      score: 0.7,
    };
    if (!materialMatches.has(material)) materialMatches.set(material, match);
  });

  const matches = [...matchesByBase.values(), ...materialMatches.values()].sort((a, b) => a.columnIndex - b.columnIndex);
  const requiredCount = REQUIRED_BASE_FIELDS.filter((field) => matchesByBase.has(field)).length;
  const baseCount = matchesByBase.size;
  const materialCount = materialMatches.size;

  if (requiredCount < 4 || baseCount < 5) return null;

  return {
    score: requiredCount * 100 + baseCount * 20 + materialCount * 5,
    baseCount,
    requiredCount,
    materialCount,
    headerRowIndex: rowIndex,
    dataStartRowIndex,
    matches,
  };
}

function detectHeader(rows: unknown[][]) {
  let best: HeaderCandidate | null = null;
  const limit = Math.min(rows.length, 8);

  for (let index = 0; index < limit; index += 1) {
    const single = evaluateHeaderCandidate(rows[index], index, index + 1);
    if (single && (!best || single.score > best.score)) best = single;

    if (index + 1 < rows.length) {
      const merged = evaluateHeaderCandidate(mergeHeaderRows(rows[index], rows[index + 1]), index, index + 2);
      if (merged && (!best || merged.score > best.score)) best = merged;
    }
  }

  return best;
}

function isEmptyRow(row: unknown[]) {
  return row.every((cell) => normalizeDisplayText(cell) === "");
}

function isRepeatedHeaderRow(row: unknown[], matches: HeaderMatch[]) {
  let repeated = 0;
  for (const match of matches) {
    const cell = normalizeLookup(row[match.columnIndex]);
    if (!cell) continue;
    if (cell === normalizeLookup(match.rawHeader) || cell === normalizeLookup(match.canonical)) repeated += 1;
  }
  return repeated >= Math.max(3, Math.floor(matches.length / 2));
}

function createEmptyBaseRecord(): ParsedImportRecord {
  return {
    QTD: "",
    AREA: "",
    UNIDADE: "",
    POSTO_GRAD: "",
    QUADRO: "",
    NOME_COMPLETO: "",
    RG: "",
    materiais: {},
  };
}

export function normalizePostoGraduacao(value: unknown) {
  const display = normalizeDisplayText(value);
  const lookup = ` ${normalizeLookup(value)} `;
  if (!display) return "";

  for (const { canonical, patterns } of POSTO_PATTERNS) {
    if (patterns.some((pattern) => lookup.includes(` ${normalizeLookup(pattern)} `))) return canonical;
  }

  return display;
}

export function normalizeMaterialValue(value: unknown) {
  const display = normalizeDisplayText(value).replace(/\s+/g, "");
  if (!display) return "";
  const lookup = stripAccents(display).replace(/[^A-Z0-9-]/g, "");
  return SIZE_MAP[lookup] ?? display;
}

export function normalizeBaseFieldValue(field: CanonicalBaseField, value: unknown) {
  const text = normalizeDisplayText(value);
  if (!text) return "";

  if (field === "POSTO_GRAD") return normalizePostoGraduacao(text);
  if (field === "RG") return stringifyCell(value).replace(/\.0+$/, "").trim();
  return text;
}

export function isBaseField(field: string): field is CanonicalBaseField {
  return (BASE_FIELDS as readonly string[]).includes(field);
}

export function getAllMaterialKeys<T extends RowRecord>(records: T[]) {
  const materials = new Set<string>();
  records.forEach((record) => Object.keys(record.materiais || {}).forEach((material) => materials.add(material)));
  return [...materials].sort((a, b) => a.localeCompare(b, "pt-BR"));
}

export function alignMaterialFields<T extends RowRecord>(records: T[]) {
  const materials = getAllMaterialKeys(records);
  return records.map((record) => ({
    ...record,
    materiais: Object.fromEntries(materials.map((material) => [material, record.materiais?.[material] ?? ""])),
  }));
}

export function dedupeRecords<T extends { RG: string; NOME_COMPLETO: string }>(records: T[]) {
  const seen = new Set<string>();
  return records.filter((record) => {
    const key = `${normalizeLookup(record.RG)}::${normalizeLookup(record.NOME_COMPLETO)}`;
    if (key === "::") return true;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export function sortRecords<T extends { AREA: string; UNIDADE: string; POSTO_GRAD: string; NOME_COMPLETO: string }>(records: T[]) {
  return [...records].sort((left, right) => {
    const unitCompare = (left.UNIDADE || left.AREA || "").localeCompare(right.UNIDADE || right.AREA || "", "pt-BR");
    if (unitCompare !== 0) return unitCompare;

    const postoCompare = (POSTO_ORDER[left.POSTO_GRAD] ?? 999) - (POSTO_ORDER[right.POSTO_GRAD] ?? 999);
    if (postoCompare !== 0) return postoCompare;

    return left.NOME_COMPLETO.localeCompare(right.NOME_COMPLETO, "pt-BR");
  });
}

export function getMaterialSizes<T extends RowRecord>(records: T[], material: string) {
  const seen = new Set<string>();
  records.forEach((record) => {
    const value = normalizeMaterialValue(record.materiais?.[material]);
    if (value) seen.add(value);
  });

  const ordered = SIZE_ORDER.filter((size) => seen.has(size));
  ordered.forEach((size) => seen.delete(size));
  return [...ordered, ...[...seen].sort((a, b) => a.localeCompare(b, "pt-BR"))];
}

function buildPreview(records: ParsedImportRecord[], materials: string[]) {
  const previewColumns = [...BASE_FIELDS.map((field) => BASE_FIELD_LABELS[field]), ...materials, "ORIGEM"];
  const previewRows = records.slice(0, 5).map((record) => ({
    QTD: record.QTD,
    "ÁREA": record.AREA,
    UNIDADE: record.UNIDADE,
    "POSTO/GRAD": record.POSTO_GRAD,
    QUADRO: record.QUADRO,
    "NOME COMPLETO": record.NOME_COMPLETO,
    RG: record.RG,
    ...Object.fromEntries(materials.map((material) => [material, record.materiais[material] ?? ""])),
    ORIGEM: "",
  }));

  return { previewColumns, previewRows };
}

function parseSheet(ws: XLSX.WorkSheet, fileName: string, sheetName: string) {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "", blankrows: false }) as unknown[][];
  if (rows.length === 0) return { records: [] as ParsedImportRecord[], warnings: [] as string[], materials: [] as string[] };

  const header = detectHeader(rows);
  if (!header) {
    return {
      records: [],
      warnings: [`Aba "${sheetName}": nenhum cabeçalho reconhecido para importação.`],
      materials: [],
    };
  }

  const warnings = new Set<string>();
  warnings.add(`Aba "${sheetName}": cabeçalho detectado na linha ${header.headerRowIndex + 1}.`);

  header.matches.forEach((match) => {
    if (normalizeLookup(match.rawHeader) !== normalizeLookup(match.canonical)) {
      warnings.add(`Aba "${sheetName}": coluna "${match.rawHeader}" mapeada para "${match.canonical}".`);
    }
  });

  const records: ParsedImportRecord[] = [];
  const materials = header.matches.filter((match) => match.kind === "material").map((match) => match.canonical);

  for (let rowIndex = header.dataStartRowIndex; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    if (isEmptyRow(row) || isRepeatedHeaderRow(row, header.matches)) continue;

    const record = createEmptyBaseRecord();

    header.matches.forEach((match) => {
      const rawValue = row[match.columnIndex];
      if (match.kind === "base" && match.field) {
        record[match.field] = normalizeBaseFieldValue(match.field, rawValue);
      }
      if (match.kind === "material") {
        record.materiais[match.canonical] = normalizeMaterialValue(rawValue);
      }
    });

    record.UNIDADE = record.UNIDADE || record.AREA;
    if (!record.NOME_COMPLETO || !record.RG) continue;
    records.push(record);
  }

  if (records.length === 0) warnings.add(`Aba "${sheetName}": nenhum registro válido encontrado.`);
  return { records, warnings: [...warnings], materials };
}

export function parseWorkbook(workbook: XLSX.WorkBook, fileName: string): ParsedWorkbookResult {
  const warnings = new Set<string>();
  const collected: ParsedImportRecord[] = [];

  workbook.SheetNames.forEach((sheetName) => {
    if (SHEETS_TO_IGNORE.some((ignored) => sheetName.toLowerCase().includes(ignored))) return;
    const result = parseSheet(workbook.Sheets[sheetName], fileName, sheetName);
    result.warnings.forEach((warning) => warnings.add(warning));
    collected.push(...result.records);
  });

  const records = alignMaterialFields(collected);
  const materials = getAllMaterialKeys(records);
  const { previewColumns, previewRows } = buildPreview(records, materials);

  return {
    records,
    warnings: [...warnings],
    materials,
    previewColumns,
    previewRows,
  };
}
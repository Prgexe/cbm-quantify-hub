import React, { createContext, useContext, useState, useCallback } from "react";
import * as XLSX from "xlsx";

export interface MilitarRecord {
  id: string;
  QTD: number | string;
  AREA: string;
  UNIDADE: string;
  POSTO_GRAD: string;
  QUADRO: string;
  NOME_COMPLETO: string;
  RG: string;
  CAMISETA_GV: string;
  CAMISA_UV: string;
  SHORT_JOHN: string;
  _source: string;
  _color: string;
}

export interface ImportedFile {
  name: string;
  recordCount: number;
  status: "success" | "error";
  error?: string;
  importedAt: Date;
}

interface DataContextType {
  records: MilitarRecord[];
  files: ImportedFile[];
  importFile: (file: File) => Promise<void>;
  removeFile: (fileName: string) => void;
  updateRecord: (id: string, field: keyof MilitarRecord, value: string) => void;
  setRecordColor: (id: string, color: string) => void;
  clearAll: () => void;
}

const DataContext = createContext<DataContextType | null>(null);

const IGNORED_SHEETS = ["resumo geral", "contagem"];

let idCounter = 0;

function norm(s: unknown): string {
  return String(s ?? "").trim().toUpperCase();
}

/**
 * Suporta dois formatos de planilha do CBMERJ:
 *
 * FORMATO A — planilha consolidada geral (ex: CONSOLIDADO_GERAL_FINAL...):
 *   row[0]: Título (nome do CBA)
 *   row[1]: vazia
 *   row[2]: QTD | ÁREA | UNIDADE (FINAL) | POSTO/GRAD | QUADRO | NOME COMPLETO | RG | Camiseta de GV | Camisa U.V para GV | "Short John"
 *   row[3+]: dados
 *
 * FORMATO B — planilha individual por quartel (ex: PLANILHA_MATERIAL_GBS):
 *   row[0]: Título longo "RELAÇÃO NOMINAL..."
 *   row[1]: QTD | ÁREA | UNIDADE (FINAL) | POSTO/GRAD | NOME COMPLETO | RG | TAMANHOS | | 
 *   row[2]: (vazios x6) | Camiseta de GV | Camisa U.V para GV | "Short John"
 *   row[3+]: dados
 */
function parseSheet(ws: XLSX.WorkSheet, fileName: string): MilitarRecord[] {
  const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "" }) as unknown[][];
  if (raw.length < 4) return [];

  // Detecta formato pela presença de "QTD" na linha index 1
  const isFormatB = raw[1]?.some((c) => norm(c) === "QTD") ?? false;

  const colMap: Record<number, keyof MilitarRecord> = {};

  function mapCell(n: string, i: number) {
    if (n === "QTD")                                         colMap[i] = "QTD";
    else if (n === "ÁREA" || n === "AREA")                   colMap[i] = "AREA";
    else if (n === "UNIDADE (FINAL)" || n === "UNIDADE")     colMap[i] = "UNIDADE";
    else if (n === "POSTO/GRAD")                             colMap[i] = "POSTO_GRAD";
    else if (n === "QUADRO")                                 colMap[i] = "QUADRO";
    else if (n === "NOME COMPLETO")                          colMap[i] = "NOME_COMPLETO";
    else if (n === "RG")                                     colMap[i] = "RG";
    else if (n === "CAMISETA DE GV" || n === "CAMISETA GV")  colMap[i] = "CAMISETA_GV";
    else if (n.startsWith("CAMISA U"))                       colMap[i] = "CAMISA_UV";
    else if (n.includes("SHORT JOHN"))                       colMap[i] = "SHORT_JOHN";
  }

  if (isFormatB) {
    // Linha 1: colunas fixas
    raw[1].forEach((cell, i) => mapCell(norm(cell), i));
    // Linha 2: colunas de tamanho
    raw[2].forEach((cell, i) => mapCell(norm(cell), i));
  } else {
    // Formato A: linha 2 tem todas as colunas
    raw[2].forEach((cell, i) => mapCell(norm(cell), i));
  }

  const records: MilitarRecord[] = [];

  for (let i = 3; i < raw.length; i++) {
    const row = raw[i];
    if (!row || row.every((c) => String(c ?? "").trim() === "")) continue;

    const rec: Partial<MilitarRecord> = {
      id: `rec_${++idCounter}`,
      _source: fileName,
      _color: "",
      QTD: "", AREA: "", UNIDADE: "", POSTO_GRAD: "", QUADRO: "",
      NOME_COMPLETO: "", RG: "", CAMISETA_GV: "", CAMISA_UV: "", SHORT_JOHN: "",
    };

    row.forEach((cell, idx) => {
      const field = colMap[idx];
      if (field) (rec as any)[field] = String(cell ?? "").trim();
    });

    if (rec.NOME_COMPLETO && rec.NOME_COMPLETO !== "") {
      records.push(rec as MilitarRecord);
    }
  }

  return records;
}

function parseWorkbook(wb: XLSX.WorkBook, fileName: string): MilitarRecord[] {
  const all: MilitarRecord[] = [];
  for (const sheetName of wb.SheetNames) {
    if (IGNORED_SHEETS.some((s) => sheetName.toLowerCase().includes(s))) continue;
    all.push(...parseSheet(wb.Sheets[sheetName], fileName));
  }
  return all;
}

export function DataProvider({ children }: { children: React.ReactNode }) {
  const [records, setRecords] = useState<MilitarRecord[]>([]);
  const [files, setFiles] = useState<ImportedFile[]>([]);

  const importFile = useCallback(async (file: File) => {
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const newRecords = parseWorkbook(wb, file.name);

      setRecords((prev) => [...prev, ...newRecords]);
      setFiles((prev) => [
        ...prev,
        { name: file.name, recordCount: newRecords.length, status: "success", importedAt: new Date() },
      ]);
    } catch (e: any) {
      setFiles((prev) => [
        ...prev,
        { name: file.name, recordCount: 0, status: "error", error: e.message, importedAt: new Date() },
      ]);
    }
  }, []);

  const removeFile = useCallback((fileName: string) => {
    setRecords((prev) => prev.filter((r) => r._source !== fileName));
    setFiles((prev) => prev.filter((f) => f.name !== fileName));
  }, []);

  const updateRecord = useCallback((id: string, field: keyof MilitarRecord, value: string) => {
    setRecords((prev) => prev.map((r) => (r.id === id ? { ...r, [field]: value } : r)));
  }, []);

  const setRecordColor = useCallback((id: string, color: string) => {
    setRecords((prev) => prev.map((r) => (r.id === id ? { ...r, _color: color } : r)));
  }, []);

  const clearAll = useCallback(() => {
    setRecords([]);
    setFiles([]);
  }, []);

  return (
    <DataContext.Provider value={{ records, files, importFile, removeFile, updateRecord, setRecordColor, clearAll }}>
      {children}
    </DataContext.Provider>
  );
}

export function useData() {
  const ctx = useContext(DataContext);
  if (!ctx) throw new Error("useData must be used within DataProvider");
  return ctx;
}

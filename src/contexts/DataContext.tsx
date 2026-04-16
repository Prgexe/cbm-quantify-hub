import React, { createContext, useCallback, useContext, useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  alignMaterialFields,
  dedupeRecords,
  isBaseField,
  normalizeBaseFieldValue,
  normalizeMaterialValue,
  parseWorkbook,
} from "@/utils/normalizer";

export interface MilitarRecord {
  id: string;
  QTD: string;
  AREA: string;
  UNIDADE: string;
  POSTO_GRAD: string;
  QUADRO: string;
  NOME_COMPLETO: string;
  RG: string;
  materiais: Record<string, string>;
  _source: string;
  _color: string;
}

export interface ImportedFile {
  name: string;
  recordCount: number;
  duplicateCount: number;
  materials: string[];
  warnings: string[];
  previewColumns: string[];
  previewRows: Array<Record<string, string>>;
  status: "success" | "error";
  error?: string;
  importedAt: Date;
}

interface DataContextType {
  records: MilitarRecord[];
  files: ImportedFile[];
  importFile: (file: File) => Promise<void>;
  removeFile: (fileName: string) => void;
  updateRecord: (id: string, field: string, value: string) => void;
  setRecordColor: (id: string, color: string) => void;
  clearAll: () => void;
}

const DataContext = createContext<DataContextType | null>(null);

let idCounter = 0;

export function DataProvider({ children }: { children: React.ReactNode }) {
  const [records, setRecords] = useState<MilitarRecord[]>([]);
  const [files, setFiles] = useState<ImportedFile[]>([]);
  const recordsRef = useRef<MilitarRecord[]>([]);

  useEffect(() => {
    recordsRef.current = records;
  }, [records]);

  const importFile = useCallback(async (file: File) => {
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const parsed = parseWorkbook(wb, file.name);
      const incoming: MilitarRecord[] = parsed.records.map((record) => ({
        ...record,
        id: `rec_${++idCounter}`,
        _source: file.name,
        _color: "",
      }));

      const existing = recordsRef.current.filter((record) => record._source !== file.name);
      const merged = alignMaterialFields(dedupeRecords([...existing, ...incoming]));
      const importedRecords = merged.filter((record) => record._source === file.name);
      const duplicateCount = incoming.length - importedRecords.length;

      if (importedRecords.length === 0) {
        throw new Error(parsed.warnings.at(-1) || "Nenhum registro válido foi encontrado nessa planilha.");
      }

      recordsRef.current = merged;
      setRecords(merged);
      setFiles((prev) => [
        ...prev.filter((current) => current.name !== file.name),
        {
          name: file.name,
          recordCount: importedRecords.length,
          duplicateCount,
          materials: parsed.materials,
          warnings: duplicateCount > 0
            ? [...parsed.warnings, `${duplicateCount} duplicata(s) removida(s) por RG + NOME COMPLETO.`]
            : parsed.warnings,
          previewColumns: parsed.previewColumns,
          previewRows: parsed.previewRows.map((row) => ({ ...row, ORIGEM: file.name })),
          status: "success",
          importedAt: new Date(),
        },
      ]);
    } catch (e: any) {
      setFiles((prev) => [
        ...prev.filter((current) => current.name !== file.name),
        {
          name: file.name,
          recordCount: 0,
          duplicateCount: 0,
          materials: [],
          warnings: [],
          previewColumns: [],
          previewRows: [],
          status: "error",
          error: e.message,
          importedAt: new Date(),
        },
      ]);
    }
  }, []);

  const removeFile = useCallback((fileName: string) => {
    const nextRecords = recordsRef.current.filter((record) => record._source !== fileName);
    recordsRef.current = nextRecords;
    setRecords(nextRecords);
    setFiles((prev) => prev.filter((f) => f.name !== fileName));
  }, []);

  const updateRecord = useCallback((id: string, field: string, value: string) => {
    const nextRecords = recordsRef.current.map((record) => {
      if (record.id !== id) return record;
      if (isBaseField(field)) {
        return { ...record, [field]: normalizeBaseFieldValue(field, value) };
      }
      return {
        ...record,
        materiais: {
          ...record.materiais,
          [field]: normalizeMaterialValue(value),
        },
      };
    });

    recordsRef.current = nextRecords;
    setRecords(nextRecords);
  }, []);

  const setRecordColor = useCallback((id: string, color: string) => {
    const nextRecords = recordsRef.current.map((record) => (record.id === id ? { ...record, _color: color } : record));
    recordsRef.current = nextRecords;
    setRecords(nextRecords);
  }, []);

  const clearAll = useCallback(() => {
    recordsRef.current = [];
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

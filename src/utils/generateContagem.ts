import * as XLSX from "xlsx";
import type { MilitarRecord } from "@/contexts/DataContext";

const SIZE_ORDER = ["PP", "P", "M", "G", "GG", "XGG", "EXG"];

const isIgnored = (v: string) => {
  const t = (v ?? "").trim();
  return t === "" || t === "--" || t === "-";
};

/**
 * Gera a worksheet "Contagem" — pivot Unidade × Material × Tamanhos.
 * Detecta materiais, tamanhos e unidades dinamicamente a partir dos dados.
 */
export function generateContagemSheet(
  records: MilitarRecord[],
  materials: string[],
): XLSX.WorkSheet {
  // 1) Tamanhos presentes globalmente, na ordem canônica
  const presentSizes = new Set<string>();
  for (const r of records) {
    for (const m of materials) {
      const v = (r.materiais?.[m] ?? "").toString().trim().toUpperCase();
      if (!isIgnored(v)) presentSizes.add(v);
    }
  }
  const sizes = SIZE_ORDER.filter((s) => presentSizes.has(s));
  // anexa quaisquer tamanhos não previstos (raro), mantendo ordem de descoberta
  for (const s of presentSizes) {
    if (!sizes.includes(s)) sizes.push(s);
  }

  // 2) Unidades únicas na ordem de aparição
  const unidades: string[] = [];
  const seen = new Set<string>();
  for (const r of records) {
    const u = r.UNIDADE || r.AREA || "Sem Unidade";
    if (!seen.has(u)) {
      seen.add(u);
      unidades.push(u);
    }
  }

  // 3) Para regra do "--": (material, tamanho) sem ocorrência alguma no dataset
  const materialHasSize: Record<string, Set<string>> = {};
  for (const m of materials) materialHasSize[m] = new Set();
  for (const r of records) {
    for (const m of materials) {
      const v = (r.materiais?.[m] ?? "").toString().trim().toUpperCase();
      if (!isIgnored(v)) materialHasSize[m].add(v);
    }
  }

  // 4) Monta linhas
  const aoa: (string | number)[][] = [];
  aoa.push(["Unidade", "Material", ...sizes, "TOTAL"]);

  const totalGeral: Record<string, Record<string, number>> = {};
  for (const m of materials) {
    totalGeral[m] = {};
    for (const s of sizes) totalGeral[m][s] = 0;
  }

  for (const unidade of unidades) {
    const recsU = records.filter(
      (r) => (r.UNIDADE || r.AREA || "Sem Unidade") === unidade,
    );
    materials.forEach((material, idx) => {
      const row: (string | number)[] = [idx === 0 ? unidade : "", material];
      let total = 0;
      for (const size of sizes) {
        if (!materialHasSize[material].has(size)) {
          row.push("--");
          continue;
        }
        const count = recsU.reduce((acc, r) => {
          const v = (r.materiais?.[material] ?? "").toString().trim().toUpperCase();
          return acc + (v === size ? 1 : 0);
        }, 0);
        row.push(count);
        total += count;
        totalGeral[material][size] += count;
      }
      row.push(total);
      aoa.push(row);
    });
  }

  // 5) Bloco TOTAL GERAL
  aoa.push([]);
  aoa.push(["TOTAL GERAL", "Material", ...sizes, "TOTAL"]);
  for (const material of materials) {
    const row: (string | number)[] = ["", material];
    let total = 0;
    for (const size of sizes) {
      if (!materialHasSize[material].has(size)) {
        row.push("--");
        continue;
      }
      const v = totalGeral[material][size];
      row.push(v);
      total += v;
    }
    row.push(total);
    aoa.push(row);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [
    { wch: 32 },
    { wch: 26 },
    ...sizes.map(() => ({ wch: 8 })),
    { wch: 10 },
  ];

  // Negrito no cabeçalho e na linha TOTAL GERAL (SheetJS community: aplica via cell.s)
  const totalGeralRow = aoa.findIndex((r) => r[0] === "TOTAL GERAL");
  const boldRows = [0, totalGeralRow].filter((i) => i >= 0);
  for (const rIdx of boldRows) {
    for (let c = 0; c < aoa[rIdx].length; c++) {
      const addr = XLSX.utils.encode_cell({ r: rIdx, c });
      const cell = ws[addr];
      if (cell) cell.s = { font: { bold: true } };
    }
  }

  return ws;
}

import * as XLSX from "xlsx";
import type { MilitarRecord } from "@/contexts/DataContext";

// Tamanhos válidos na ordem de exibição
const SIZE_ORDER = ["PP", "P", "M", "G", "GG", "XG"];

const isIgnored = (v: string) => {
  const t = (v ?? "").trim();
  return t === "" || t === "--" || t === "-" || t === "X" || t === "XX" || t === "XXX";
};

// ─── Estilos ────────────────────────────────────────────────────────────────

const HEADER_STYLE = {
  font: { bold: true, color: { rgb: "FFFFFF" }, name: "Arial", sz: 10 },
  fill: { fgColor: { rgb: "1F3864" }, patternType: "solid" },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
  border: {
    top: { style: "thin", color: { rgb: "888888" } }, bottom: { style: "thin", color: { rgb: "888888" } },
    left: { style: "thin", color: { rgb: "888888" } }, right: { style: "thin", color: { rgb: "888888" } },
  },
};

const TOTAL_GERAL_STYLE = {
  font: { bold: true, color: { rgb: "FFFFFF" }, name: "Arial", sz: 10 },
  fill: { fgColor: { rgb: "1D4E2A" }, patternType: "solid" },
  alignment: { horizontal: "center", vertical: "center" },
  border: {
    top: { style: "thin", color: { rgb: "888888" } }, bottom: { style: "thin", color: { rgb: "888888" } },
    left: { style: "thin", color: { rgb: "888888" } }, right: { style: "thin", color: { rgb: "888888" } },
  },
};

const UNIT_STYLE = {
  font: { bold: true, name: "Arial", sz: 10 },
  fill: { fgColor: { rgb: "D6E4F0" }, patternType: "solid" },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
  border: {
    top: { style: "medium", color: { rgb: "555555" } }, bottom: { style: "medium", color: { rgb: "555555" } },
    left: { style: "medium", color: { rgb: "555555" } }, right: { style: "medium", color: { rgb: "555555" } },
  },
};

const cellBorder = {
  top: { style: "thin", color: { rgb: "CCCCCC" } }, bottom: { style: "thin", color: { rgb: "CCCCCC" } },
  left: { style: "thin", color: { rgb: "CCCCCC" } }, right: { style: "thin", color: { rgb: "CCCCCC" } },
};

const makeCellStyle = (isAlt: boolean, bold = false) => ({
  font: { name: "Arial", sz: 10, bold },
  fill: isAlt ? { fgColor: { rgb: "EBF0FA" }, patternType: "solid" } : {},
  alignment: { horizontal: "center", vertical: "center" },
  border: cellBorder,
});

const makeMaterialStyle = (isAlt: boolean, bold = false) => ({
  font: { name: "Arial", sz: 10, bold },
  fill: isAlt ? { fgColor: { rgb: "EBF0FA" }, patternType: "solid" } : {},
  alignment: { horizontal: "left", vertical: "center" },
  border: cellBorder,
});

const makeTotalStyle = (isAlt: boolean) => ({
  font: { name: "Arial", sz: 10, bold: true },
  fill: { fgColor: { rgb: isAlt ? "C6EFCE" : "E2EFDA" }, patternType: "solid" },
  alignment: { horizontal: "center", vertical: "center" },
  border: cellBorder,
});

// ─── Helper: define célula com fórmula na worksheet ─────────────────────────

function setFormula(ws: XLSX.WorkSheet, addr: string, formula: string, style?: object) {
  ws[addr] = { t: "n", f: formula, v: 0 };
  if (style) (ws[addr] as any).s = style;
}

function setString(ws: XLSX.WorkSheet, addr: string, value: string, style?: object) {
  ws[addr] = { t: "s", v: value };
  if (style) (ws[addr] as any).s = style;
}

function setNumber(ws: XLSX.WorkSheet, addr: string, value: number, style?: object) {
  ws[addr] = { t: "n", v: value };
  if (style) (ws[addr] as any).s = style;
}

// ─── Função principal ────────────────────────────────────────────────────────

/**
 * Gera a worksheet "Contagem" com fórmulas COUNTIFS que referenciam
 * a aba "Consolidado Geral". Assim qualquer edição nos dados atualiza
 * automaticamente os totais.
 *
 * Estrutura da aba Consolidado Geral (gerada pelo exportXlsx):
 *   Col A = ÁREA, Col B = UNIDADE, Col C = POSTO/GRAD, Col D = QUADRO,
 *   Col E = NOME COMPLETO, Col F = RG, Col G+ = materiais
 */
export function generateContagemSheet(
  records: MilitarRecord[],
  materials: string[],
): XLSX.WorkSheet {

  // ── 1. Detecta tamanhos válidos presentes nos dados ──────────────────────
  const presentSizes = new Set<string>();
  for (const r of records) {
    for (const m of materials) {
      const v = (r.materiais?.[m] ?? "").toString().trim().toUpperCase();
      if (!isIgnored(v)) presentSizes.add(v);
    }
  }
  const sizes = SIZE_ORDER.filter((s) => presentSizes.has(s));
  for (const s of presentSizes) {
    if (!sizes.includes(s)) sizes.push(s);
  }

  // ── 2. Unidades únicas na ordem de aparição ──────────────────────────────
  const unidades: string[] = [];
  const seen = new Set<string>();
  for (const r of records) {
    const u = r.UNIDADE || r.AREA || "Sem Unidade";
    if (!seen.has(u)) { seen.add(u); unidades.push(u); }
  }

  // ── 3. Detecta quais (material × tamanho) existem no dataset ─────────────
  // Para saber quando colocar "--" em vez de fórmula
  const materialHasSize: Record<string, Set<string>> = {};
  for (const m of materials) materialHasSize[m] = new Set();
  for (const r of records) {
    for (const m of materials) {
      const v = (r.materiais?.[m] ?? "").toString().trim().toUpperCase();
      if (!isIgnored(v)) materialHasSize[m].add(v);
    }
  }

  // ── 4. Descobre em qual coluna do Consolidado Geral cada material está ───
  // O toRow() no ConsolidatedTab exporta nesta ordem:
  //   ÁREA(A), UNIDADE(B), POSTO/GRAD(C), QUADRO(D), NOME COMPLETO(E), RG(F), materiais(G+)
  const BASE_COL_COUNT = 6; // A até F
  const materialColLetter: Record<string, string> = {};
  materials.forEach((m, i) => {
    materialColLetter[m] = XLSX.utils.encode_col(BASE_COL_COUNT + i); // G, H, I...
  });

  // Referência à aba Consolidado Geral (com aspas simples para nomes com espaço)
  const SRC = "'Consolidado Geral'";

  // Linha máxima dos dados no Consolidado Geral (linha 1 = cabeçalho, dados a partir da 2)
  const dataRows = records.length;
  const lastDataRow = dataRows + 1; // +1 pelo cabeçalho

  // Coluna UNIDADE no Consolidado Geral = B
  const UNIDADE_COL = "B";

  // ── 5. Monta a worksheet ─────────────────────────────────────────────────
  const ws: XLSX.WorkSheet = { "!type": "sheet" };
  const merges: XLSX.Range[] = [];
  const numCols = 2 + sizes.length + 1; // Unidade + Material + tamanhos + TOTAL

  // Linha de início dos dados (0-indexed para o código, 1-indexed nas fórmulas Excel)
  let currentRow = 0; // 0-indexed

  // Cabeçalho
  const hdr = ["Unidade", "Material", ...sizes, "TOTAL"];
  hdr.forEach((h, c) => {
    setString(ws, XLSX.utils.encode_cell({ r: currentRow, c }), h, HEADER_STYLE);
  });
  currentRow++;

  // Blocos por unidade
  for (const unidade of unidades) {
    const blockStart = currentRow;

    materials.forEach((material, mIdx) => {
      const isAlt = currentRow % 2 === 0;
      const excelRow = currentRow + 1; // 1-indexed para fórmulas

      // Col 0: Unidade (só preenche na primeira linha do bloco; o restante é mesclado)
      setString(ws, XLSX.utils.encode_cell({ r: currentRow, c: 0 }),
        mIdx === 0 ? unidade : "", UNIT_STYLE);

      // Col 1: Material
      setString(ws, XLSX.utils.encode_cell({ r: currentRow, c: 1 }),
        material, makeMaterialStyle(isAlt));

      // Colunas de tamanho: fórmula COUNTIFS
      const matCol = materialColLetter[material]; // ex: "G"
      const totalFormulaParts: string[] = [];

      sizes.forEach((size, sIdx) => {
        const c = 2 + sIdx;
        const addr = XLSX.utils.encode_cell({ r: currentRow, c });

        if (!materialHasSize[material].has(size)) {
          // Tamanho não existe para este material → "--"
          setString(ws, addr, "--", makeCellStyle(isAlt));
        } else {
          // COUNTIFS(UNIDADE_col, unidade, MATERIAL_col, tamanho)
          const formula =
            `COUNTIFS(${SRC}!${UNIDADE_COL}$2:${UNIDADE_COL}$${lastDataRow},"${unidade}",` +
            `${SRC}!${matCol}$2:${matCol}$${lastDataRow},"${size}")`;
          setFormula(ws, addr, formula, makeCellStyle(isAlt));
          totalFormulaParts.push(XLSX.utils.encode_cell({ r: currentRow, c }));
        }
      });

      // Col TOTAL: soma das colunas de tamanho desta linha
      const totalCol = 2 + sizes.length;
      const totalAddr = XLSX.utils.encode_cell({ r: currentRow, c: totalCol });
      if (totalFormulaParts.length > 0) {
        setFormula(ws, totalAddr,
          `SUM(${XLSX.utils.encode_cell({ r: currentRow, c: 2 })}:${XLSX.utils.encode_cell({ r: currentRow, c: totalCol - 1 })})`,
          makeTotalStyle(isAlt));
      } else {
        setNumber(ws, totalAddr, 0, makeTotalStyle(isAlt));
      }

      currentRow++;
    });

    const blockEnd = currentRow - 1;

    // Mescla coluna Unidade para todo o bloco
    if (blockEnd > blockStart) {
      merges.push({ s: { r: blockStart, c: 0 }, e: { r: blockEnd, c: 0 } });
    }
  }

  // ── Linha vazia de separação ─────────────────────────────────────────────
  currentRow++;
  const totalGeralHeaderRow = currentRow;

  // ── Cabeçalho TOTAL GERAL ────────────────────────────────────────────────
  const tgHdr = ["TOTAL GERAL", "Material", ...sizes, "TOTAL"];
  tgHdr.forEach((h, c) => {
    setString(ws, XLSX.utils.encode_cell({ r: currentRow, c }), h, TOTAL_GERAL_STYLE);
  });
  currentRow++;

  const totalGeralStart = currentRow;

  // ── Linhas de TOTAL GERAL: COUNTIFS sem filtro de unidade ───────────────
  materials.forEach((material) => {
    const isAlt = currentRow % 2 === 0;
    const matCol = materialColLetter[material];

    setString(ws, XLSX.utils.encode_cell({ r: currentRow, c: 0 }), "", TOTAL_GERAL_STYLE);
    setString(ws, XLSX.utils.encode_cell({ r: currentRow, c: 1 }), material,
      makeMaterialStyle(isAlt, true));

    sizes.forEach((size, sIdx) => {
      const c = 2 + sIdx;
      const addr = XLSX.utils.encode_cell({ r: currentRow, c });

      if (!materialHasSize[material].has(size)) {
        setString(ws, addr, "--", makeCellStyle(isAlt, true));
      } else {
        // COUNTIF simples: conta em todo o dataset
        const formula =
          `COUNTIF(${SRC}!${matCol}$2:${matCol}$${lastDataRow},"${size}")`;
        setFormula(ws, addr, formula, makeCellStyle(isAlt, true));
      }
    });

    const totalCol = 2 + sizes.length;
    setFormula(ws, XLSX.utils.encode_cell({ r: currentRow, c: totalCol }),
      `SUM(${XLSX.utils.encode_cell({ r: currentRow, c: 2 })}:${XLSX.utils.encode_cell({ r: currentRow, c: totalCol - 1 })})`,
      makeTotalStyle(isAlt));

    currentRow++;
  });

  const totalGeralEnd = currentRow - 1;

  // Mescla coluna 0 do bloco TOTAL GERAL
  if (totalGeralEnd > totalGeralStart) {
    merges.push({
      s: { r: totalGeralHeaderRow, c: 0 },
      e: { r: totalGeralEnd, c: 0 },
    });
  }

  // ── Metadados da sheet ───────────────────────────────────────────────────
  const lastCell = XLSX.utils.encode_cell({ r: currentRow - 1, c: numCols - 1 });
  ws["!ref"] = `A1:${lastCell}`;

  ws["!cols"] = [
    { wch: 34 },
    { wch: 26 },
    ...sizes.map(() => ({ wch: 8 })),
    { wch: 10 },
  ];

  ws["!merges"] = merges;

  const lastColLetter = XLSX.utils.encode_col(numCols - 1);
  ws["!autofilter"] = { ref: `A1:${lastColLetter}1` };

  return ws;
}

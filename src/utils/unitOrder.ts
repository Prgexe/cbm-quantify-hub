/**
 * unitOrder.ts
 * Define a ordem hierárquica das unidades dentro de cada CBA.
 *
 * Para adicionar novas unidades: insira na lista do CBA correto.
 * Unidades não listadas aparecem no final em ordem alfabética.
 *
 * A comparação é feita com normUnit() que ignora acentos, pontuação,
 * ordinal (°/º), barras e espaços — portanto nomes antigos como
 * "1º GMAR", "DBM1/M", "DBM 5/M SEPETIBA" são reconhecidos corretamente.
 */

const UNIT_ORDER: Record<string, string[]> = {

  // ── CBA V ──────────────────────────────────────────────────────────────────
  "CBA V": [
    "CBA V - BAIXADAS LITORÂNEAS",
    // 9º GBM e subordinados
    "9º GBM - MACAÉ",
    "DBM 2/9 - RIO DAS OSTRAS",
    // 18º GBM e subordinados
    "18º GBM – CABO FRIO",
    "DBM 1/18 - SÃO PEDRO DA ALDEIA",
    "DBM 2/18 - ARMAÇÃO DOS BÚZIOS",
    "PABM 1/18 - ARRAIAL DO CABO",
    // 27º GBM e subordinados
    "27º GBM - ARARUAMA",
    "DBM 1/27 – SAQUAREMA",
  ],

  // ── CBA VI ─────────────────────────────────────────────────────────────────
  "CBA VI": [
    "CBA VI",
  ],

  // ── CBA VII ────────────────────────────────────────────────────────────────
  "CBA VII": [
    "26º GBM - PARATY",
    "DBM 1/26 - PARATY",   // cobre: "DBM 1/26", "DBM1/26"
  ],

  // ── CBA VIII ───────────────────────────────────────────────────────────────
  "CBA VIII": [
    "CBA VIII - ATIVIDADES ESPECIALIZADAS",
    "2º GSFMA",
    "GBRESC",
    "GOPP",
    "COVANT",
    "GRUPAMENTO TÉCNICO DE SUPRIMENTO DE ÁGUA PARA INCÊNDIO - GTSAI",
  ],

  // ── CBA X ──────────────────────────────────────────────────────────────────
  "CBA X": [
    "CBA X",
    // 1º GMAR e subordinados
    "1º GMAR - Botafogo",          // cobre: "1º GMAR", "1 GMAR", "1°GMAR"
    "DBM 1/M - Paquetá",           // cobre: "DBM1/M", "DBM 1/M", "DBM1M"
    "DBM 2/M - Piscinão de Ramos", // cobre: "DBM 2/M", "DBM2/M"
    // 2º GMAR e subordinados
    "2º GMAR - Barra da Tijuca",   // cobre: "2º GMAR", "2 GMAR", "2°GMAR"
    "DBM 3/M - Recreio dos Bandeirantes",
    "DBM 4/M - Barra de Guaratiba", // cobre: "DBM 4/M", "DBM4/M"
    "DBM 5/M - Sepetiba",           // cobre: "DBM 5/M SEPETIBA", "DBM5/M"
    // 3º GMAR
    "3º GMAR - Copacabana",         // cobre: "3º GMAR", "3 GMAR"
    // 4º GMAR
    "4º GMAR - Itaipu",             // cobre: "4º GMAR", "4 GMAR"
    // DBM 1/26 (subordinado ao 26º GBM mas registrado como CBA X)
    "DBM 1/26",
  ],

};

// ─── Normalização ─────────────────────────────────────────────────────────────

/**
 * Normaliza string para comparação robusta:
 * - Remove acentos e diacríticos
 * - Remove ordinal (° e º)
 * - Substitui toda pontuação/barra/hífen por espaço
 * - Separa letras coladas a dígitos (DBM1 → DBM 1)
 * - Converte para MAIÚSCULAS e colapsa espaços
 *
 * Resultado: "DBM1/M" → "DBM 1 M", igual a "DBM 1/M - Paquetá"
 */
function normUnit(s: string): string {
  let r = s.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // remove acentos
  r = r.replace(/[°º]/g, "");                                  // remove ordinal
  r = r.replace(/[^A-Za-z0-9]/g, " ");                        // pontuação → espaço
  r = r.replace(/(?<=[A-Za-z])(?=[0-9])/g, " ");              // DBM1 → DBM 1
  r = r.replace(/(?<=[0-9])(?=[A-Za-z])/g, " ");              // 1M → 1 M
  return r.replace(/\s+/g, " ").trim().toUpperCase();
}

// ─── Índice de ordenação ──────────────────────────────────────────────────────

function getUnitIndex(area: string, unidade: string): number {
  const order = UNIT_ORDER[area];
  if (!order) return 9999;

  const normU = normUnit(unidade);

  // 1. Correspondência exata normalizada
  const exact = order.findIndex((u) => normUnit(u) === normU);
  if (exact !== -1) return exact;

  // 2. Correspondência por prefixo
  //    "1 GMAR" casa com "1 GMAR BOTAFOGO"
  //    "DBM 5 M SEPETIBA" casa com "DBM 5 M SEPETIBA" (já exact) ou variações
  const prefix = order.findIndex((u) => {
    const nu = normUnit(u);
    return nu.startsWith(normU + " ") || normU.startsWith(nu + " ");
  });
  if (prefix !== -1) return prefix;

  return 9999;
}

// ─── Exportações ──────────────────────────────────────────────────────────────

/**
 * Ordena registros pela hierarquia de unidades definida em UNIT_ORDER.
 * Dentro de cada unidade, mantém a ordem relativa original.
 */
export function sortByUnitHierarchy<T extends { AREA: string; UNIDADE: string }>(
  records: T[]
): T[] {
  return [...records].sort((a, b) => {
    const idxA = getUnitIndex(a.AREA, a.UNIDADE);
    const idxB = getUnitIndex(b.AREA, b.UNIDADE);
    if (idxA !== idxB) return idxA - idxB;
    return normUnit(a.UNIDADE).localeCompare(normUnit(b.UNIDADE));
  });
}

/**
 * Retorna as unidades únicas de um array de registros
 * já na ordem hierárquica correta.
 */
export function getOrderedUnidades<T extends { AREA: string; UNIDADE: string }>(
  records: T[]
): string[] {
  const seen = new Set<string>();
  const sorted = sortByUnitHierarchy(records);
  const result: string[] = [];
  for (const r of sorted) {
    if (!seen.has(r.UNIDADE)) {
      seen.add(r.UNIDADE);
      result.push(r.UNIDADE);
    }
  }
  return result;
}

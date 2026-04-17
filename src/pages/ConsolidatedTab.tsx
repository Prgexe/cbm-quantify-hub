import { useMemo, useState } from "react";
import { Download, Search, Maximize2, Minimize2 } from "lucide-react";
import * as XLSX from "xlsx";
import { useData, MilitarRecord } from "@/contexts/DataContext";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { BASE_FIELDS, BASE_FIELD_LABELS, getAllMaterialKeys, isBaseField, sortRecords } from "@/utils/normalizer";
import { generateContagemSheet } from "@/utils/generateContagem";

const HIGHLIGHT_COLORS = [
  { value: "",       label: "— Sem cor",     bg: "" },
  { value: "yellow", label: "🟡 Amarelo",    bg: "hsl(var(--row-yellow))" },
  { value: "green",  label: "🟢 Verde",      bg: "hsl(var(--row-green))" },
  { value: "red",    label: "🔴 Vermelho",   bg: "hsl(var(--row-red))" },
  { value: "blue",   label: "🔵 Azul",       bg: "hsl(var(--row-blue))" },
  { value: "orange", label: "🟠 Laranja",    bg: "hsl(var(--row-orange))" },
  { value: "purple", label: "🟣 Roxo",       bg: "hsl(var(--secondary))" },
];

const COLOR_BG: Record<string, string> = Object.fromEntries(
  HIGHLIGHT_COLORS.map((c) => [c.value, c.bg])
);

const PAGE_SIZE = 100;

const WIDTHS: Record<string, string> = {
  AREA: "200px", UNIDADE: "300px", POSTO_GRAD: "160px",
  QUADRO: "90px", NOME_COMPLETO: "280px", RG: "70px",
  ORIGEM: "220px",
};

interface ColumnDef {
  id: string;
  label: string;
  width: string;
  editable: boolean;
}

export default function ConsolidatedTab() {
  const { records, updateRecord, setRecordColor } = useData();
  const [search, setSearch]             = useState("");
  const [filterArea, setFilterArea]     = useState("__all__");
  const [filterUnidade, setFilterUnidade] = useState("__all__");
  const [filterPosto, setFilterPosto]   = useState("__all__");
  const [filterQuadro, setFilterQuadro] = useState("__all__");
  const [page, setPage]                 = useState(0);
  const [editingCell, setEditingCell]   = useState<{ id: string; field: string } | null>(null);
  const [fullscreen, setFullscreen]     = useState(false);
  const materials = useMemo(() => getAllMaterialKeys(records), [records]);
  const columns = useMemo<ColumnDef[]>(() => [
    ...BASE_FIELDS.map((field) => ({ id: field, label: BASE_FIELD_LABELS[field], width: WIDTHS[field] || "140px", editable: true })),
    ...materials.map((material) => ({ id: material, label: material, width: WIDTHS[material] || "110px", editable: true })),
  ], [materials]);

  const areas    = useMemo(() => [...new Set(records.map((r) => r.AREA))].filter(Boolean).sort(), [records]);
  const unidades = useMemo(() => [...new Set(records.map((r) => r.UNIDADE))].filter(Boolean).sort(), [records]);
  const postos   = useMemo(() => [...new Set(records.map((r) => r.POSTO_GRAD))].filter(Boolean).sort(), [records]);
  const quadros  = useMemo(() => [...new Set(records.map((r) => r.QUADRO))].filter(Boolean).sort(), [records]);

  const filtered = useMemo(() => {
    let data = records;
    if (search) {
      const s = search.toLowerCase();
      data = data.filter((r) =>
        r.NOME_COMPLETO.toLowerCase().includes(s) || r.RG.toLowerCase().includes(s)
      );
    }
    if (filterArea    !== "__all__") data = data.filter((r) => r.AREA      === filterArea);
    if (filterUnidade !== "__all__") data = data.filter((r) => r.UNIDADE   === filterUnidade);
    if (filterPosto   !== "__all__") data = data.filter((r) => r.POSTO_GRAD === filterPosto);
    if (filterQuadro  !== "__all__") data = data.filter((r) => r.QUADRO    === filterQuadro);
    return sortRecords(data);
  }, [records, search, filterArea, filterUnidade, filterPosto, filterQuadro]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const currentPage = Math.min(page, totalPages - 1);
  const paged = filtered.slice(currentPage * PAGE_SIZE, (currentPage + 1) * PAGE_SIZE);

  // ── Exportação: aba por CBA + aba geral + aba Contagem (formato pivot) ──
  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();

    const toRow = (record: MilitarRecord) => ({
      "ÁREA": record.AREA,
      UNIDADE: record.UNIDADE,
      "POSTO/GRAD": record.POSTO_GRAD,
      QUADRO: record.QUADRO,
      "NOME COMPLETO": record.NOME_COMPLETO,
      RG: record.RG,
      ...Object.fromEntries(materials.map((material) => [material, record.materiais[material] ?? ""])),
    });

    const colWidths = columns
      .filter((c) => c.id !== "ORIGEM")
      .map((column) => ({ wch: Math.max(10, Math.round(parseInt(column.width, 10) / 8) || 14) }));

    // Agrupa por CBA/AREA já normalizado pelo importador
    const grouped: Record<string, MilitarRecord[]> = {};
    for (const r of filtered) {
      const key = r.AREA || "Sem Área";
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(r);
    }

    // Uma aba por CBA
    for (const [area, rows] of Object.entries(grouped).sort()) {
      const ws = XLSX.utils.json_to_sheet(rows.map(toRow));
      ws["!cols"] = colWidths;
      XLSX.utils.book_append_sheet(wb, ws, area.substring(0, 31));
    }

    // Aba consolidada geral
    const wsAll = XLSX.utils.json_to_sheet(filtered.map(toRow));
    wsAll["!cols"] = colWidths;
    XLSX.utils.book_append_sheet(wb, wsAll, "Consolidado Geral");

    // ── Aba CONTAGEM (formato pivot: Unidade × Material × Tamanhos) ──
    const SIZE_COLS = ["PP", "P", "M", "G", "GG", "XG", "X"];
    const headerSizes = SIZE_COLS.filter((sz) =>
      filtered.some((r) => materials.some((m) => (r.materiais[m] || "").toUpperCase() === sz))
    );

    // Unidades únicas, ordenadas
    const unidadesUnicas = [...new Set(filtered.map((r) => r.UNIDADE || r.AREA || "Sem Unidade"))].sort();

    const contagem: any[][] = [];
    contagem.push(["Unidade", "Material", ...headerSizes, "TOTAL"]);

    const totalGeral: Record<string, Record<string, number>> = {};
    materials.forEach((m) => {
      totalGeral[m] = {};
      headerSizes.forEach((s) => (totalGeral[m][s] = 0));
      totalGeral[m].TOTAL = 0;
    });

    unidadesUnicas.forEach((unidade) => {
      const recsUnidade = filtered.filter((r) => (r.UNIDADE || r.AREA) === unidade);
      materials.forEach((material) => {
        const row: any[] = [unidade, material];
        let totalLinha = 0;
        headerSizes.forEach((size) => {
          const count = recsUnidade.reduce(
            (acc, r) => acc + (((r.materiais[material] || "").toUpperCase() === size) ? 1 : 0),
            0
          );
          row.push(count || (size === "PP" || size === "X" ? "-" : 0));
          if (typeof count === "number") {
            totalLinha += count;
            totalGeral[material][size] += count;
          }
        });
        row.push(totalLinha);
        totalGeral[material].TOTAL += totalLinha;
        contagem.push(row);
      });
    });

    // Linha em branco + Total Geral
    contagem.push([]);
    contagem.push(["TOTAL GERAL", "Material", ...headerSizes, "TOTAL"]);
    materials.forEach((material) => {
      const row: any[] = ["", material];
      headerSizes.forEach((size) => {
        const v = totalGeral[material][size];
        row.push(v || (size === "PP" || size === "X" ? "-" : 0));
      });
      row.push(totalGeral[material].TOTAL);
      contagem.push(row);
    });

    const wsContagem = XLSX.utils.aoa_to_sheet(contagem);
    wsContagem["!cols"] = [
      { wch: 32 },
      { wch: 22 },
      ...headerSizes.map(() => ({ wch: 8 })),
      { wch: 10 },
    ];

    XLSX.utils.book_append_sheet(wb, wsContagem, "Contagem");

    XLSX.writeFile(wb, "consolidado_cbmerj.xlsx");
  };

  const wrapClass = fullscreen
    ? "fixed inset-0 z-50 bg-background flex flex-col p-4 overflow-hidden"
    : "space-y-4";

  return (
    <div className={wrapClass}>
      {/* Cabeçalho */}
      <div className="flex flex-wrap items-center justify-between gap-3">
        <h2 className="text-xl font-bold">📋 Planilha Consolidada</h2>
        <div className="flex gap-2">
          <Button variant="outline" size="sm" onClick={() => setFullscreen((f) => !f)}>
            {fullscreen ? <Minimize2 className="w-4 h-4 mr-1" /> : <Maximize2 className="w-4 h-4 mr-1" />}
            {fullscreen ? "Sair" : "Tela cheia"}
          </Button>
          <Button onClick={exportXlsx} className="bg-accent text-accent-foreground hover:bg-accent/90">
            <Download className="w-4 h-4 mr-2" /> Exportar XLSX
          </Button>
        </div>
      </div>

      {/* Filtros */}
      <div className="flex flex-wrap gap-2">
        <div className="relative flex-1 min-w-[200px]">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
          <Input
            placeholder="Buscar por nome ou RG..."
            value={search}
            onChange={(e) => { setSearch(e.target.value); setPage(0); }}
            className="pl-9"
          />
        </div>
        <FSel label="Área"    value={filterArea}    options={areas}    onChange={(v) => { setFilterArea(v);    setPage(0); }} />
        <FSel label="Unidade" value={filterUnidade} options={unidades} onChange={(v) => { setFilterUnidade(v); setPage(0); }} />
        <FSel label="Posto"   value={filterPosto}   options={postos}   onChange={(v) => { setFilterPosto(v);   setPage(0); }} />
        <FSel label="Quadro"  value={filterQuadro}  options={quadros}  onChange={(v) => { setFilterQuadro(v);  setPage(0); }} />
      </div>

      <p className="text-sm text-muted-foreground">{filtered.length} registros encontrados</p>

      {/* Tabela */}
      <div className={`overflow-auto border rounded-lg ${fullscreen ? "flex-1" : "max-h-[70vh]"}`}>
        <table className="text-sm border-collapse" style={{ minWidth: "1500px" }}>
          <thead className="sticky top-0 z-10">
            <tr className="bg-primary text-primary-foreground">
              <th className="p-2 text-left text-xs whitespace-nowrap sticky left-0 bg-primary z-20" style={{ minWidth: "130px" }}>
                Marcação
              </th>
              {columns.map((column) => (
                <th key={column.id} className="p-2 text-left text-xs whitespace-nowrap" style={{ minWidth: column.width }}>
                  {column.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {paged.map((r) => {
              const bg = COLOR_BG[r._color] || "";
              return (
                <tr key={r.id} className="border-t" style={{ backgroundColor: bg || undefined }}>
                  {/* Marcação */}
                  <td className="p-1 sticky left-0 z-10 border-r" style={{ backgroundColor: bg || "white" }}>
                    <select
                      value={r._color}
                      onChange={(e) => setRecordColor(r.id, e.target.value)}
                      className="w-full rounded border border-input bg-card px-1 py-0.5 text-xs"
                    >
                      {HIGHLIGHT_COLORS.map((c) => (
                        <option key={c.value} value={c.value}>{c.label}</option>
                      ))}
                    </select>
                  </td>
                  {/* Campos */}
                  {columns.map((column) => (
                    <ECell
                      key={column.id}
                      r={r} field={column.id} editable={column.editable} bg={bg}
                      editing={editingCell?.id === r.id && editingCell?.field === column.id}
                      onStartEdit={() => column.editable && setEditingCell({ id: r.id, field: column.id })}
                      onEndEdit={(v) => { updateRecord(r.id, column.id, v); setEditingCell(null); }}
                      onCancel={() => setEditingCell(null)}
                    />
                  ))}
                </tr>
              );
            })}
            {paged.length === 0 && (
              <tr>
                <td colSpan={columns.length + 1} className="p-8 text-center text-muted-foreground">
                  Nenhum registro encontrado.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Paginação */}
      <div className="flex items-center justify-between pt-1">
        <Button variant="outline" size="sm" disabled={page === 0} onClick={() => setPage(page - 1)}>
          Anterior
        </Button>
        <span className="text-sm text-muted-foreground">
          Página {currentPage + 1} de {totalPages} · {filtered.length} registros
        </span>
        <Button variant="outline" size="sm" disabled={currentPage >= totalPages - 1} onClick={() => setPage(currentPage + 1)}>
          Próxima
        </Button>
      </div>
    </div>
  );
}

function getCellValue(record: MilitarRecord, field: string) {
  if (field === "ORIGEM") return record._source;
  if (isBaseField(field)) return record[field];
  return record.materiais[field] ?? "";
}

function ECell({ r, field, editable, bg, editing, onStartEdit, onEndEdit, onCancel }: {
  r: MilitarRecord; field: string; editable: boolean; bg: string;
  editing: boolean;
  onStartEdit: () => void;
  onEndEdit: (v: string) => void;
  onCancel: () => void;
}) {
  const value = getCellValue(r, field);

  return (
    <td
      className={`p-2 whitespace-nowrap ${editable ? "cursor-pointer" : "cursor-default"}`}
      style={{ backgroundColor: bg || undefined }}
      onClick={onStartEdit}
    >
      {editing ? (
        <input
          autoFocus
          defaultValue={String(value ?? "")}
          onBlur={(e) => onEndEdit(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") (e.target as HTMLInputElement).blur();
            if (e.key === "Escape") onCancel();
          }}
          className="w-full rounded border border-input bg-card px-1 py-0.5 text-sm"
          style={{ minWidth: "60px" }}
          onClick={(e) => e.stopPropagation()}
        />
      ) : (
        <span>{String(value ?? "")}</span>
      )}
    </td>
  );
}

function FSel({ label, value, options, onChange }: {
  label: string; value: string; options: string[]; onChange: (v: string) => void;
}) {
  return (
    <Select value={value} onValueChange={onChange}>
      <SelectTrigger className="w-[160px]">
        <SelectValue placeholder={label} />
      </SelectTrigger>
      <SelectContent>
        <SelectItem value="__all__">Todos ({label})</SelectItem>
        {options.map((o) => <SelectItem key={o} value={o}>{o}</SelectItem>)}
      </SelectContent>
    </Select>
  );
}

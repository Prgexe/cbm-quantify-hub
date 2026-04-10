import { useMemo, useState } from "react";
import { Download, Search, Maximize2, Minimize2 } from "lucide-react";
import * as XLSX from "xlsx";
import { useData, MilitarRecord } from "@/contexts/DataContext";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

const HIGHLIGHT_COLORS = [
  { value: "",       label: "— Sem cor",     bg: "" },
  { value: "yellow", label: "🟡 Amarelo",    bg: "#FEF08A" },
  { value: "green",  label: "🟢 Verde",      bg: "#BBF7D0" },
  { value: "red",    label: "🔴 Vermelho",   bg: "#FECACA" },
  { value: "blue",   label: "🔵 Azul",       bg: "#BFDBFE" },
  { value: "orange", label: "🟠 Laranja",    bg: "#FED7AA" },
  { value: "purple", label: "🟣 Roxo",       bg: "#E9D5FF" },
];

const COLOR_BG: Record<string, string> = Object.fromEntries(
  HIGHLIGHT_COLORS.map((c) => [c.value, c.bg])
);

const PAGE_SIZE = 100;

const FIELDS: (keyof MilitarRecord)[] = [
  "QTD", "AREA", "UNIDADE", "POSTO_GRAD", "QUADRO",
  "NOME_COMPLETO", "RG", "CAMISETA_GV", "CAMISA_UV", "SHORT_JOHN",
];

const LABELS: Record<string, string> = {
  QTD: "QTD", AREA: "ÁREA", UNIDADE: "UNIDADE", POSTO_GRAD: "POSTO/GRAD",
  QUADRO: "QUADRO", NOME_COMPLETO: "NOME COMPLETO", RG: "RG",
  CAMISETA_GV: "Camiseta GV", CAMISA_UV: "Camisa U.V", SHORT_JOHN: "Short John",
};

const WIDTHS: Record<string, string> = {
  QTD: "50px", AREA: "200px", UNIDADE: "300px", POSTO_GRAD: "130px",
  QUADRO: "90px", NOME_COMPLETO: "280px", RG: "70px",
  CAMISETA_GV: "90px", CAMISA_UV: "90px", SHORT_JOHN: "90px",
};

export default function ConsolidatedTab() {
  const { records, updateRecord, setRecordColor } = useData();
  const [search, setSearch]             = useState("");
  const [filterArea, setFilterArea]     = useState("__all__");
  const [filterUnidade, setFilterUnidade] = useState("__all__");
  const [filterPosto, setFilterPosto]   = useState("__all__");
  const [filterQuadro, setFilterQuadro] = useState("__all__");
  const [page, setPage]                 = useState(0);
  const [editingCell, setEditingCell]   = useState<{ id: string; field: keyof MilitarRecord } | null>(null);
  const [fullscreen, setFullscreen]     = useState(false);

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
    return data;
  }, [records, search, filterArea, filterUnidade, filterPosto, filterQuadro]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const paged = filtered.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);

  // ── Exportação: uma aba por CBA + aba geral ──────────────────────────────────
  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();

    const toRow = (r: MilitarRecord) => ({
      "QTD": r.QTD,
      "ÁREA": r.AREA,
      "UNIDADE (FINAL)": r.UNIDADE,
      "POSTO/GRAD": r.POSTO_GRAD,
      "QUADRO": r.QUADRO,
      "NOME COMPLETO": r.NOME_COMPLETO,
      "RG": r.RG,
      "Camiseta de GV": r.CAMISETA_GV,
      "Camisa U.V para GV": r.CAMISA_UV,
      '"Short John"': r.SHORT_JOHN,
    });

    const colWidths = [
      { wch: 5 }, { wch: 25 }, { wch: 35 }, { wch: 15 }, { wch: 10 },
      { wch: 35 }, { wch: 8 }, { wch: 14 }, { wch: 18 }, { wch: 12 },
    ];

    // Normaliza área: travessão → hífen para evitar abas duplicadas
    const normalizeArea = (a: string) =>
      (a || "Sem Área").replace(/[–—]/g, "-").replace(/\s+/g, " ").trim();

    // Agrupa por CBA/AREA normalizado
    const grouped: Record<string, MilitarRecord[]> = {};
    for (const r of filtered) {
      const key = normalizeArea(r.AREA);
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
              {FIELDS.map((f) => (
                <th key={f} className="p-2 text-left text-xs whitespace-nowrap" style={{ minWidth: WIDTHS[f] }}>
                  {LABELS[f]}
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
                      className="w-full text-xs rounded border px-1 py-0.5 bg-white"
                    >
                      {HIGHLIGHT_COLORS.map((c) => (
                        <option key={c.value} value={c.value}>{c.label}</option>
                      ))}
                    </select>
                  </td>
                  {/* Campos */}
                  {FIELDS.map((field) => (
                    <ECell
                      key={field}
                      r={r} field={field} bg={bg}
                      editing={editingCell?.id === r.id && editingCell?.field === field}
                      onStartEdit={() => setEditingCell({ id: r.id, field })}
                      onEndEdit={(v) => { updateRecord(r.id, field, v); setEditingCell(null); }}
                      onCancel={() => setEditingCell(null)}
                    />
                  ))}
                </tr>
              );
            })}
            {paged.length === 0 && (
              <tr>
                <td colSpan={12} className="p-8 text-center text-muted-foreground">
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
          Página {page + 1} de {totalPages} · {filtered.length} registros
        </span>
        <Button variant="outline" size="sm" disabled={page >= totalPages - 1} onClick={() => setPage(page + 1)}>
          Próxima
        </Button>
      </div>
    </div>
  );
}

function ECell({ r, field, bg, editing, onStartEdit, onEndEdit, onCancel }: {
  r: MilitarRecord; field: keyof MilitarRecord; bg: string;
  editing: boolean;
  onStartEdit: () => void;
  onEndEdit: (v: string) => void;
  onCancel: () => void;
}) {
  return (
    <td
      className="p-2 cursor-pointer whitespace-nowrap"
      style={{ backgroundColor: bg || undefined }}
      onClick={onStartEdit}
    >
      {editing ? (
        <input
          autoFocus
          defaultValue={String(r[field] ?? "")}
          onBlur={(e) => onEndEdit(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") (e.target as HTMLInputElement).blur();
            if (e.key === "Escape") onCancel();
          }}
          className="w-full px-1 py-0.5 border rounded text-sm bg-white"
          style={{ minWidth: "60px" }}
          onClick={(e) => e.stopPropagation()}
        />
      ) : (
        <span>{String(r[field] ?? "")}</span>
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

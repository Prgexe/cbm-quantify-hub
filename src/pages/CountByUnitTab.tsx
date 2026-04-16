import { useMemo } from "react";
import * as XLSX from "xlsx";
import { Download } from "lucide-react";
import { useData } from "@/contexts/DataContext";
import { Button } from "@/components/ui/button";
import type { MilitarRecord } from "@/contexts/DataContext";
import { getAllMaterialKeys, getMaterialSizes, normalizeMaterialValue } from "@/utils/normalizer";

function countSizes(values: string[], sizes: string[]): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const s of sizes) counts[s] = 0;

  for (const raw of values) {
    const v = normalizeMaterialValue(raw) || "--";
    if (v in counts) {
      counts[v]++;
    } else {
      // Tamanho desconhecido cai em "--" se existir no mapa
      if ("--" in counts) counts["--"]++;
    }
  }
  return counts;
}

interface TableRow {
  unit: string;
  counts: Record<string, number>;
  total: number;
}

interface TableData {
  rows: TableRow[];
  totals: Record<string, number>;
  grandTotal: number;
}

function buildTable(
  records: MilitarRecord[],
  material: string,
  sizes: string[]
): TableData {
  // Agrupa por unidade
  const grouped: Record<string, MilitarRecord[]> = {};
  for (const r of records) {
    const key = r.UNIDADE || r.AREA || "(Sem Unidade)";
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(r);
  }

  const rows: TableRow[] = Object.keys(grouped)
    .sort()
    .map((unit) => {
      const values = grouped[unit].map((r) => r.materiais[material] ?? "");
      const counts = countSizes(values, sizes);
      const total = Object.values(counts).reduce((a, b) => a + b, 0);
      return { unit, counts, total };
    });

  const totals: Record<string, number> = {};
  for (const s of sizes) {
    totals[s] = rows.reduce((acc, r) => acc + (r.counts[s] ?? 0), 0);
  }
  const grandTotal = rows.reduce((acc, r) => acc + r.total, 0);

  return { rows, totals, grandTotal };
}

export default function CountByUnitTab() {
  const { records } = useData();
  const materials = useMemo(() => getAllMaterialKeys(records), [records]);
  const tables = useMemo(
    () => materials.map((material) => {
      const sizes = getMaterialSizes(records, material);
      return { material, sizes: sizes.length ? sizes : ["--"], data: buildTable(records, material, sizes.length ? sizes : ["--"]) };
    }),
    [materials, records],
  );

  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();

    const addSheet = (name: string, data: TableData, sizes: string[]) => {
      const rows = data.rows.map((r) => {
        const row: Record<string, string | number> = { UNIDADE: r.unit };
        for (const s of sizes) row[s] = r.counts[s] ?? 0;
        row["TOTAL"] = r.total;
        return row;
      });
      const totalRow: Record<string, string | number> = { UNIDADE: "TOTAL GERAL" };
      for (const s of sizes) totalRow[s] = data.totals[s] ?? 0;
      totalRow["TOTAL"] = data.grandTotal;
      rows.push(totalRow);

      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, name);
    };

    tables.forEach(({ material, data, sizes }) => addSheet(material.substring(0, 31), data, sizes));
    XLSX.writeFile(wb, "contagem_por_unidade.xlsx");
  };

  if (records.length === 0 || materials.length === 0) {
    return (
      <div className="text-center py-12 text-muted-foreground">
        Importe dados para ver a contagem por unidade.
      </div>
    );
  }

  return (
    <div className="space-y-8">
      <div className="flex items-center justify-between">
        <h2 className="text-xl font-bold">📊 Contagem por Unidade</h2>
        <Button onClick={exportXlsx} className="bg-accent text-accent-foreground hover:bg-accent/90">
          <Download className="w-4 h-4 mr-2" /> Exportar XLSX
        </Button>
      </div>

      {tables.map(({ material, data, sizes }) => (
        <SummaryTable key={material} title={material} data={data} sizes={sizes} />
      ))}
    </div>
  );
}

function SummaryTable({
  title,
  data,
  sizes,
}: {
  title: string;
  data: TableData;
  sizes: string[];
}) {
  return (
    <div>
      <h3 className="font-semibold mb-2">{title}</h3>
      <div className="overflow-x-auto border rounded-lg">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-primary text-primary-foreground">
              <th className="p-2 text-left">UNIDADE</th>
              {sizes.map((s) => (
                <th key={s} className="p-2 text-center">
                  {s}
                </th>
              ))}
              <th className="p-2 text-center font-bold">TOTAL</th>
            </tr>
          </thead>
          <tbody>
            {data.rows.map((r) => (
              <tr key={r.unit} className="border-t hover:bg-muted/30">
                <td className="p-2 whitespace-nowrap">{r.unit}</td>
                {sizes.map((s) => (
                  <td key={s} className="p-2 text-center">
                    {r.counts[s] ?? 0}
                  </td>
                ))}
                <td className="p-2 text-center font-semibold">{r.total}</td>
              </tr>
            ))}
            <tr className="border-t bg-muted font-bold">
              <td className="p-2">TOTAL GERAL</td>
              {sizes.map((s) => (
                <td key={s} className="p-2 text-center">
                  {data.totals[s] ?? 0}
                </td>
              ))}
              <td className="p-2 text-center">{data.grandTotal}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

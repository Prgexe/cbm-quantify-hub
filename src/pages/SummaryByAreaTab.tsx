import { useMemo } from "react";
import * as XLSX from "xlsx";
import { Download, Users } from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";
import { useData } from "@/contexts/DataContext";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";

const CAMISETA_SIZES = ["PP", "P", "M", "G", "GG", "XGG", "EXG", "--"];
const CAMISA_SIZES = ["P", "M", "G", "GG", "XGG", "--"];
const SHORT_SIZES = ["P", "M", "G", "GG", "XGG", "--"];

function countByGroup(records: ReturnType<typeof useData>["records"], groupField: "AREA", itemField: "CAMISETA_GV" | "CAMISA_UV" | "SHORT_JOHN", sizes: string[]) {
  const grouped: Record<string, Record<string, number>> = {};
  for (const r of records) {
    const group = r[groupField] || "Sem Área";
    if (!grouped[group]) {
      grouped[group] = {};
      for (const s of sizes) grouped[group][s] = 0;
    }
    const v = r[itemField].trim().toUpperCase() || "--";
    if (v in grouped[group]) grouped[group][v]++;
    else grouped[group]["--"]++;
  }

  const rows = Object.keys(grouped).sort().map((group) => {
    const total = Object.values(grouped[group]).reduce((a, b) => a + b, 0);
    return { group, counts: grouped[group], total };
  });

  const totals: Record<string, number> = {};
  for (const s of sizes) totals[s] = rows.reduce((a, r) => a + (r.counts[s] || 0), 0);
  const grandTotal = rows.reduce((a, r) => a + r.total, 0);

  return { rows, totals, grandTotal };
}

const CHART_COLORS = ["#1e3a5f", "#c62828", "#2e7d32", "#e65100", "#1565c0", "#6a1b9a", "#00838f"];

export default function SummaryByAreaTab() {
  const { records } = useData();

  const areaCount = useMemo(() => {
    const map: Record<string, number> = {};
    for (const r of records) {
      const a = r.AREA || "Sem Área";
      map[a] = (map[a] || 0) + 1;
    }
    return Object.entries(map).sort((a, b) => a[0].localeCompare(b[0])).map(([area, count]) => ({ area, count }));
  }, [records]);

  const camiseta = useMemo(() => countByGroup(records, "AREA", "CAMISETA_GV", CAMISETA_SIZES), [records]);
  const camisa = useMemo(() => countByGroup(records, "AREA", "CAMISA_UV", CAMISA_SIZES), [records]);
  const shortJohn = useMemo(() => countByGroup(records, "AREA", "SHORT_JOHN", SHORT_SIZES), [records]);

  const chartData = useMemo(() => {
    return camiseta.rows.map((r) => ({ name: r.group.substring(0, 15), ...r.counts }));
  }, [camiseta]);

  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();
    const addSheet = (name: string, data: typeof camiseta, sizes: string[]) => {
      const rows = data.rows.map((r) => ({ ÁREA: r.group, ...r.counts, TOTAL: r.total }));
      rows.push({ ÁREA: "TOTAL GERAL", ...data.totals, TOTAL: data.grandTotal });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), name);
    };
    addSheet("Camiseta GV por Área", camiseta, CAMISETA_SIZES);
    addSheet("Camisa UV por Área", camisa, CAMISA_SIZES);
    addSheet("Short John por Área", shortJohn, SHORT_SIZES);
    XLSX.writeFile(wb, "resumo_por_area.xlsx");
  };

  if (records.length === 0) {
    return <div className="text-center py-12 text-muted-foreground">Importe dados para ver o resumo por área.</div>;
  }

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <h2 className="text-xl font-bold">🗂️ Resumo por Área (CBA)</h2>
        <Button onClick={exportXlsx} className="bg-accent text-accent-foreground hover:bg-accent/90">
          <Download className="w-4 h-4 mr-2" /> Exportar XLSX
        </Button>
      </div>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        {areaCount.map((a) => (
          <Card key={a.area}>
            <CardContent className="p-4 flex items-center gap-3">
              <Users className="w-5 h-5 text-accent" />
              <div>
                <p className="text-xs text-muted-foreground">{a.area}</p>
                <p className="text-lg font-bold">{a.count}</p>
              </div>
            </CardContent>
          </Card>
        ))}
      </div>

      {chartData.length > 0 && (
        <Card>
          <CardHeader><CardTitle className="text-base">Distribuição de Camiseta GV por Área</CardTitle></CardHeader>
          <CardContent>
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis />
                <Tooltip />
                <Legend />
                {CAMISETA_SIZES.filter((s) => s !== "--").map((s, i) => (
                  <Bar key={s} dataKey={s} fill={CHART_COLORS[i % CHART_COLORS.length]} stackId="a" />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>
      )}

      <GroupTable title="Camiseta de GV por Área" data={camiseta} sizes={CAMISETA_SIZES} />
      <GroupTable title="Camisa U.V por Área" data={camisa} sizes={CAMISA_SIZES} />
      <GroupTable title="Short John por Área" data={shortJohn} sizes={SHORT_SIZES} />
    </div>
  );
}

function GroupTable({ title, data, sizes }: { title: string; data: ReturnType<typeof countByGroup>; sizes: string[] }) {
  return (
    <div>
      <h3 className="font-semibold mb-2">{title}</h3>
      <div className="overflow-x-auto border rounded-lg">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-primary text-primary-foreground">
              <th className="p-2 text-left">ÁREA</th>
              {sizes.map((s) => <th key={s} className="p-2 text-center">{s}</th>)}
              <th className="p-2 text-center font-bold">TOTAL</th>
            </tr>
          </thead>
          <tbody>
            {data.rows.map((r) => (
              <tr key={r.group} className="border-t hover:bg-muted/30">
                <td className="p-2 whitespace-nowrap">{r.group}</td>
                {sizes.map((s) => <td key={s} className="p-2 text-center">{r.counts[s] || 0}</td>)}
                <td className="p-2 text-center font-semibold">{r.total}</td>
              </tr>
            ))}
            <tr className="border-t bg-muted font-bold">
              <td className="p-2">TOTAL GERAL</td>
              {sizes.map((s) => <td key={s} className="p-2 text-center">{data.totals[s]}</td>)}
              <td className="p-2 text-center">{data.grandTotal}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

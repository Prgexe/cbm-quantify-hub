import { Trash2, FileSpreadsheet, AlertTriangle } from "lucide-react";
import { useData } from "@/contexts/DataContext";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";

export default function SettingsTab() {
  const { files, clearAll, records } = useData();

  return (
    <div className="space-y-6 max-w-2xl">
      <h2 className="text-xl font-bold">⚙️ Configurações</h2>

      <Card>
        <CardHeader><CardTitle className="text-base">Dados Importados</CardTitle></CardHeader>
        <CardContent className="space-y-3">
          <p className="text-sm text-muted-foreground">
            Total de registros: <span className="font-semibold text-foreground">{records.length}</span> | 
            Arquivos: <span className="font-semibold text-foreground">{files.length}</span>
          </p>

          {files.length > 0 && (
            <div className="space-y-2">
              {files.map((f) => (
                <div key={f.name} className="flex items-center gap-3 p-2 bg-muted/50 rounded text-sm">
                  <FileSpreadsheet className="w-4 h-4 text-success" />
                  <span className="flex-1">{f.name}</span>
                  <span className="text-xs text-muted-foreground">{f.recordCount} reg. — {f.importedAt.toLocaleString("pt-BR")}</span>
                </div>
              ))}
            </div>
          )}
        </CardContent>
      </Card>

      <Card className="border-destructive/30">
        <CardHeader><CardTitle className="text-base text-destructive flex items-center gap-2"><AlertTriangle className="w-4 h-4" /> Zona de Perigo</CardTitle></CardHeader>
        <CardContent>
          <p className="text-sm text-muted-foreground mb-3">Limpar todos os dados importados. Esta ação não pode ser desfeita.</p>
          <Button variant="destructive" onClick={clearAll} disabled={records.length === 0}>
            <Trash2 className="w-4 h-4 mr-2" /> Limpar Todos os Dados
          </Button>
        </CardContent>
      </Card>
    </div>
  );
}

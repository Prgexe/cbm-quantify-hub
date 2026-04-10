import { useCallback, useRef, useState } from "react";
import { Upload, FileSpreadsheet, Trash2, CheckCircle2, AlertTriangle } from "lucide-react";
import { useData } from "@/contexts/DataContext";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";

export default function UploadTab() {
  const { importFile, files, removeFile } = useData();
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFiles = useCallback(
    async (fileList: FileList) => {
      for (const file of Array.from(fileList)) {
        if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls")) {
          await importFile(file);
        }
      }
    },
    [importFile]
  );

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      if (e.dataTransfer.files) handleFiles(e.dataTransfer.files);
    },
    [handleFiles]
  );

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-bold">📥 Upload & Importação</h2>
        <p className="text-sm text-muted-foreground">Envie as planilhas .xlsx de cada quartel para consolidação.</p>
      </div>

      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
        onClick={() => inputRef.current?.click()}
        className={`border-2 border-dashed rounded-lg p-12 text-center cursor-pointer transition-colors ${
          dragging ? "border-accent bg-accent/5" : "border-border hover:border-accent/50"
        }`}
      >
        <Upload className="w-10 h-10 mx-auto mb-3 text-muted-foreground" />
        <p className="font-medium">Arraste arquivos .xlsx aqui ou clique para selecionar</p>
        <p className="text-sm text-muted-foreground mt-1">Cada arquivo representa um quartel/CBA</p>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls"
          multiple
          className="hidden"
          onChange={(e) => e.target.files && handleFiles(e.target.files)}
        />
      </div>

      {files.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle className="text-base">Arquivos Importados ({files.length})</CardTitle>
          </CardHeader>
          <CardContent className="space-y-2">
            {files.map((f) => (
              <div key={f.name} className="flex items-center justify-between p-3 rounded-md bg-muted/50">
                <div className="flex items-center gap-3">
                  <FileSpreadsheet className="w-5 h-5 text-success" />
                  <div>
                    <p className="font-medium text-sm">{f.name}</p>
                    <p className="text-xs text-muted-foreground">
                      {f.status === "success" ? `${f.recordCount} registros` : f.error}
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  {f.status === "success" ? (
                    <CheckCircle2 className="w-4 h-4 text-success" />
                  ) : (
                    <AlertTriangle className="w-4 h-4 text-warning" />
                  )}
                  <Button variant="ghost" size="icon" onClick={() => removeFile(f.name)}>
                    <Trash2 className="w-4 h-4 text-destructive" />
                  </Button>
                </div>
              </div>
            ))}
          </CardContent>
        </Card>
      )}
    </div>
  );
}

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
          <CardContent className="space-y-4">
            {files.map((f) => (
              <div key={f.name} className="rounded-lg border bg-muted/30 p-4 space-y-3">
                <div className="flex items-start justify-between gap-3">
                  <div className="flex items-start gap-3">
                    <FileSpreadsheet className="w-5 h-5 mt-0.5 text-success" />
                    <div className="space-y-1">
                      <p className="font-medium text-sm break-all">{f.name}</p>
                      <p className="text-xs text-muted-foreground">
                        {f.status === "success"
                          ? `${f.recordCount} registros válidos${f.duplicateCount ? ` · ${f.duplicateCount} duplicatas removidas` : ""}`
                          : f.error}
                      </p>
                      {f.status === "success" && f.materials.length > 0 && (
                        <div className="flex flex-wrap gap-1 pt-1">
                          {f.materials.map((material) => (
                            <span key={material} className="rounded-full bg-secondary px-2 py-0.5 text-[11px] font-medium text-secondary-foreground">
                              {material}
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
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

                {f.status === "success" && f.warnings.length > 0 && (
                  <div className="rounded-md border bg-background p-3">
                    <p className="text-xs font-semibold mb-2">Avisos de normalização</p>
                    <ul className="space-y-1 text-xs text-muted-foreground">
                      {f.warnings.map((warning, index) => (
                        <li key={`${f.name}-warning-${index}`}>• {warning}</li>
                      ))}
                    </ul>
                  </div>
                )}

                {f.status === "success" && f.previewRows.length > 0 && (
                  <div className="space-y-2">
                    <p className="text-xs font-semibold">Preview da consolidação</p>
                    <div className="overflow-x-auto rounded-md border bg-background">
                      <table className="w-full text-xs">
                        <thead>
                          <tr className="bg-primary text-primary-foreground">
                            {f.previewColumns.map((column) => (
                              <th key={`${f.name}-${column}`} className="p-2 text-left whitespace-nowrap">
                                {column}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {f.previewRows.map((row, rowIndex) => (
                            <tr key={`${f.name}-row-${rowIndex}`} className="border-t">
                              {f.previewColumns.map((column) => (
                                <td key={`${f.name}-${rowIndex}-${column}`} className="p-2 whitespace-nowrap">
                                  {row[column] || "—"}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </CardContent>
        </Card>
      )}
    </div>
  );
}

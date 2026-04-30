import { useState, useRef, useCallback } from "react";
import { Upload, GitMerge, Download, FileSpreadsheet } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

const API_URL = import.meta.env.VITE_API_URL || "https://cbm-quantify-hub.onrender.com";

interface SheetInfo {
  name: string;
  rows: number;
  unidades: string[];
}

type Step = "idle" | "loading_info" | "ready" | "merging" | "done" | "error";

const IGNORED = ["Resumo Geral", "Contagem", "Acrescentar1"];

export default function MergeTab() {
  const [consolidadaFile, setConsolidadaFile] = useState<File | null>(null);
  const [individualFile,  setIndividualFile]  = useState<File | null>(null);
  const [sheets,          setSheets]          = useState<SheetInfo[]>([]);
  const [abaDestino,      setAbaDestino]      = useState("");
  const [inserirAntes,    setInserirAntes]     = useState("");
  const [inserirModo,    setInserirModo]      = useState<"antes"|"depois">("depois");
  const [step,            setStep]            = useState<Step>("idle");
  const [error,           setError]           = useState("");
  const [downloadUrl,     setDownloadUrl]      = useState("");
  const [downloadName,    setDownloadName]     = useState("");

  const consRef = useRef<HTMLInputElement>(null);
  const indRef  = useRef<HTMLInputElement>(null);

  const onConsolidadaChange = useCallback(async (file: File) => {
    setConsolidadaFile(file);
    setSheets([]);
    setAbaDestino("");
    setInserirAntes("");
    setInserirModo("depois");
    setStep("loading_info");
    setError("");
    try {
      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(`${API_URL}/info`, { method: "POST", body: fd });
      if (!res.ok) throw new Error(await res.text());
      const data = await res.json();
      // Garante que unidades é sempre um array de strings
      const sheets: SheetInfo[] = (data.sheets || []).map((s: any) => ({
        name: String(s.name || ""),
        rows: Number(s.rows || 0),
        unidades: Array.isArray(s.unidades)
          ? s.unidades.filter((u: any) => u != null && String(u).trim() !== "")
          : [],
      }));
      setSheets(sheets);
      setStep("ready");
    } catch (e: any) {
      setError(`Erro ao ler planilha: ${e.message}`);
      setStep("error");
    }
  }, []);

  // Abas visíveis (exclui abas de resumo)
  const visibleSheets = sheets.filter(s => !IGNORED.includes(s.name));

  // Aba selecionada
  const selectedSheet = visibleSheets.find(s => s.name === abaDestino) ?? null;

  const handleMerge = async () => {
    const abaDestinoAtual = abaDestino.trim();
    const inserirAntesAtual = inserirAntes === "__fim__" ? "" : inserirAntes;
    if (!consolidadaFile || !individualFile || !abaDestinoAtual) {
      setError("Selecione a aba de destino antes de mesclar.");
      return;
    }
    setStep("merging");
    setError("");
    try {
      const fd = new FormData();
      fd.append("consolidada",      consolidadaFile);
      fd.append("individual",       individualFile);
      fd.append("aba_destino",      abaDestinoAtual);
      fd.append("inserir_antes_de", inserirAntesAtual);
      fd.append("inserir_modo",      inserirModo);

      const res = await fetch(`${API_URL}/merge`, { method: "POST", body: fd });
      if (!res.ok) throw new Error(await res.text());

      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const name = (consolidadaFile.name || "consolidado").replace(".xlsx", "_atualizado.xlsx");
      setDownloadUrl(url);
      setDownloadName(name);
      setStep("done");
    } catch (e: any) {
      setError(`Erro ao mesclar: ${e.message}`);
      setStep("ready");
    }
  };

  const reset = () => {
    setConsolidadaFile(null);
    setIndividualFile(null);
    setSheets([]);
    setAbaDestino("");
    setInserirAntes("");
    setInserirModo("depois");
    setStep("idle");
    setError("");
    setDownloadUrl("");
    setDownloadName("");
    if (consRef.current) consRef.current.value = "";
    if (indRef.current)  indRef.current.value  = "";
  };

  return (
    <div className="space-y-6 max-w-2xl">
      <div>
        <h2 className="text-xl font-bold">🔀 Mesclar Planilhas</h2>
        <p className="text-sm text-muted-foreground mt-1">
          Insira militares de uma planilha individual dentro da planilha consolidada,
          preservando formatação, cores e atualizando a contagem automaticamente.
        </p>
      </div>

      {/* Passo 1 */}
      <Card>
        <CardHeader>
          <CardTitle className="text-sm font-semibold">1. Planilha Consolidada (base)</CardTitle>
        </CardHeader>
        <CardContent>
          <DropZone
            label="Arraste ou clique para selecionar a planilha consolidada"
            file={consolidadaFile}
            inputRef={consRef}
            onChange={onConsolidadaChange}
          />
          {step === "loading_info" && (
            <p className="text-xs text-muted-foreground mt-2 animate-pulse">Lendo abas...</p>
          )}
        </CardContent>
      </Card>

      {/* Passo 2 */}
      {visibleSheets.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle className="text-sm font-semibold">2. Planilha Individual (novos militares)</CardTitle>
          </CardHeader>
          <CardContent>
            <DropZone
              label="Arraste ou clique para selecionar a planilha do quartel"
              file={individualFile}
              inputRef={indRef}
              onChange={setIndividualFile}
            />
          </CardContent>
        </Card>
      )}

      {/* Passo 3 */}
      {visibleSheets.length > 0 && individualFile && (
        <Card>
          <CardHeader>
            <CardTitle className="text-sm font-semibold">3. Configurar inserção</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            {/* Aba destino */}
            <div>
              <label className="text-xs font-medium mb-1 block">Aba da planilha consolidada</label>
              <Select value={abaDestino} onValueChange={(v) => { setAbaDestino(v); setInserirAntes(""); }}>
                <SelectTrigger>
                  <SelectValue placeholder="Selecione a aba..." />
                </SelectTrigger>
                <SelectContent>
                  {visibleSheets.map(s => (
                    <SelectItem key={s.name} value={s.name}>
                      {s.name} ({s.rows} linhas)
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            {/* Posição de inserção */}
            {selectedSheet && (
              <div className="space-y-3">
                <div>
                  <label className="text-xs font-medium mb-1 block">Referência de unidade</label>
                  <Select
                    value={inserirAntes || "__fim__"}
                    onValueChange={(v) => setInserirAntes(v === "__fim__" ? "" : v)}
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="Selecione a unidade..." />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__fim__">— Adicionar no final da aba —</SelectItem>
                      {selectedSheet.unidades.map((u, i) => (
                        <SelectItem key={`${u}-${i}`} value={u}>{u}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                {inserirAntes && (
                  <div>
                    <label className="text-xs font-medium mb-2 block">Inserir em relação à unidade</label>
                    <div className="flex gap-2">
                      <button
                        type="button"
                        onClick={() => setInserirModo("antes")}
                        className={`flex-1 py-2 px-3 rounded-md text-sm font-medium border transition-colors ${
                          inserirModo === "antes"
                            ? "bg-primary text-primary-foreground border-primary"
                            : "bg-background border-input hover:bg-muted"
                        }`}
                      >
                        ↑ Inserir antes
                      </button>
                      <button
                        type="button"
                        onClick={() => setInserirModo("depois")}
                        className={`flex-1 py-2 px-3 rounded-md text-sm font-medium border transition-colors ${
                          inserirModo === "depois"
                            ? "bg-primary text-primary-foreground border-primary"
                            : "bg-background border-input hover:bg-muted"
                        }`}
                      >
                        ↓ Inserir depois
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}
          </CardContent>
        </Card>
      )}

      {/* Erro */}
      {error && (
        <div className="p-3 rounded-md bg-destructive/10 text-destructive text-sm border border-destructive/20">
          {error}
        </div>
      )}

      {/* Botão mesclar */}
      {visibleSheets.length > 0 && individualFile && abaDestino && step !== "done" && (
        <Button
          onClick={handleMerge}
          disabled={step === "merging"}
          className="w-full bg-accent text-accent-foreground hover:bg-accent/90"
          size="lg"
        >
          {step === "merging"
            ? <span className="animate-pulse">Mesclando...</span>
            : <><GitMerge className="w-4 h-4 mr-2" />Mesclar e gerar planilha</>
          }
        </Button>
      )}

      {/* Download */}
      {step === "done" && downloadUrl && (
        <Card className="border-green-200 bg-green-50">
          <CardContent className="pt-4 space-y-3">
            <p className="text-sm font-medium text-green-800">✅ Planilha gerada com sucesso!</p>
            <p className="text-xs text-green-700">
              Os militares foram inseridos, formatação preservada e contagem atualizada.
            </p>
            <div className="flex gap-2">
              <a href={downloadUrl} download={downloadName} className="flex-1">
                <Button className="w-full" size="sm">
                  <Download className="w-4 h-4 mr-2" /> Baixar {downloadName}
                </Button>
              </a>
              <Button variant="outline" size="sm" onClick={reset}>
                Nova mesclagem
              </Button>
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function DropZone({ label, file, inputRef, onChange }: {
  label: string;
  file: File | null;
  inputRef: React.RefObject<HTMLInputElement>;
  onChange: (f: File) => void;
}) {
  const [dragging, setDragging] = useState(false);

  return (
    <div
      className={`border-2 border-dashed rounded-lg p-6 text-center cursor-pointer transition-colors ${
        dragging ? "border-accent bg-accent/5" : "border-border hover:border-accent/50"
      } ${file ? "border-green-400 bg-green-50" : ""}`}
      onDragOver={e => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => {
        e.preventDefault();
        setDragging(false);
        const f = e.dataTransfer.files[0];
        if (f) onChange(f);
      }}
      onClick={() => inputRef.current?.click()}
    >
      {file ? (
        <div className="flex items-center justify-center gap-2 text-green-700">
          <FileSpreadsheet className="w-5 h-5" />
          <span className="text-sm font-medium">{file.name}</span>
        </div>
      ) : (
        <>
          <Upload className="w-6 h-6 mx-auto mb-2 text-muted-foreground" />
          <p className="text-sm text-muted-foreground">{label}</p>
        </>
      )}
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        className="hidden"
        onChange={e => { const f = e.target.files?.[0]; if (f) onChange(f); }}
      />
    </div>
  );
}

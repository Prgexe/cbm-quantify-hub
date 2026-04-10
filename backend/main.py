from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from copy import copy
import io
import re
import zipfile
from collections import Counter
from typing import List

app = FastAPI(title="CBMERJ Almoxarifado API")

import sys, logging
logging.basicConfig(level=logging.INFO)
_log = logging.getLogger(__name__)
_log.info(f"Python: {sys.version}")
_log.info(f"openpyxl: {openpyxl.__version__}")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def repair_xlsx(data: bytes) -> bytes:
    """
    Repara xlsx gerado por LibreOffice/Google Sheets ou com XML invalido.
    - Remove caracteres de controle ilegais
    - Remove arquivos nao suportados pelo openpyxl (docProps/custom.xml)
    - Limpa referencia do custom.xml no Content_Types.xml
    """
    buf = io.BytesIO(data)
    zin = zipfile.ZipFile(buf, "r")
    names = zin.namelist()

    SKIP = {"docProps/custom.xml"}

    # First pass: build repaired zip
    out = io.BytesIO()
    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        if item.filename in SKIP:
            continue
        raw = zin.read(item.filename)
        if item.filename.endswith(".xml") or item.filename.endswith(".rels"):
            raw_str = raw.decode("utf-8", errors="replace")
            raw_str = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", raw_str)
            # Remove custom.xml override from Content_Types
            if item.filename == "[Content_Types].xml":
                raw_str = re.sub(r'<Override[^>]*custom\.xml[^/]*/>', '', raw_str)
            raw = raw_str.encode("utf-8")
        zout.writestr(item, raw)
    zout.close()
    out.seek(0)
    return out.read()


def load_workbook_safe(data: bytes, read_only: bool = False, data_only: bool = False):
    """
    Tenta carregar o workbook normalmente; se falhar, repara e tenta novamente.
    Suporta arquivos do Excel, LibreOffice e Google Sheets.
    """
    import traceback, logging
    logger = logging.getLogger(__name__)

    # Attempt 1: direct load
    try:
        return openpyxl.load_workbook(
            io.BytesIO(data), read_only=read_only, data_only=data_only
        )
    except Exception as first_err:
        logger.error(f"Attempt 1 failed: {first_err}")

    # Attempt 2: repair XML then load
    try:
        repaired = repair_xlsx(data)
        return openpyxl.load_workbook(
            io.BytesIO(repaired), read_only=read_only, data_only=data_only
        )
    except Exception as second_err:
        logger.error(f"Attempt 2 failed: {second_err}")
        logger.error(traceback.format_exc())
        raise HTTPException(
            400,
            f"Nao foi possivel ler o arquivo Excel. Verifique se nao esta corrompido "
            f"ou salve-o novamente pelo Excel. Detalhe: {first_err}"
        )


# ── Helpers ──────────────────────────────────────────────────────────────────

def make_fill(hex_color: str) -> PatternFill:
    rgb = hex_color[-6:] if len(hex_color) > 6 else hex_color
    return PatternFill(fill_type="solid", fgColor=rgb)

def get_color(ws, row_idx: int) -> str:
    try:
        fill = ws.cell(row=row_idx, column=1).fill
        if fill.fill_type is None or fill.fill_type == "none":
            return "FFFFFF"
        rgb = fill.fgColor.rgb
        # 00000000 means transparent/no color, treat as white
        if rgb in ("00000000", "000000"):
            return "FFFFFF"
        return rgb[-6:] if len(rgb) > 6 else rgb
    except:
        return "FFFFFF"

POSTO_MAP = {
    # Coronel
    "CEL": "CORONEL", "COR": "CORONEL",
    # Tenente Coronel
    "TEM CEL": "TENENTE CORONEL", "TEN CEL": "TENENTE CORONEL",
    "TC": "TENENTE CORONEL", "T CEL": "TENENTE CORONEL",
    "TENENTE-CORONEL": "TENENTE CORONEL",
    # Major
    "MAJ": "MAJOR",
    # Capitao
    "CAP": "CAPITAO",
    # Tenente
    "TEN": "TENENTE",
    "1 TEN": "1 TENENTE", "1TEN": "1 TENENTE",
    "2 TEN": "2 TENENTE", "2TEN": "2 TENENTE",
    # Subtenente
    "SUBTEN": "SUBTENENTE", "ST": "SUBTENENTE", "SUB TEN": "SUBTENENTE",
    # Sargento
    "1 SGT": "1 SARGENTO", "1SGT": "1 SARGENTO",
    "2 SGT": "2 SARGENTO", "2SGT": "2 SARGENTO",
    "3 SGT": "3 SARGENTO", "3SGT": "3 SARGENTO",
    # Cabo / Soldado
    "CB": "CABO", "SD": "SOLDADO", "SOL": "SOLDADO",
}

def normalizar_posto(posto):
    if not posto:
        return posto
    key = str(posto).strip().upper()
    return POSTO_MAP.get(key, posto)


def normalizar_area(area: str, areas_consolidada: list) -> str:
    if not area or not areas_consolidada:
        return area
    normalized_upper = area.strip().upper()
    for a in areas_consolidada:
        if a.strip().upper() == normalized_upper:
            return a
    for a in areas_consolidada:
        if a.strip().upper().startswith(normalized_upper):
            return a
    for a in areas_consolidada:
        if normalized_upper in a.strip().upper():
            return a
    return area.strip()


# ── Leitura de planilha individual ───────────────────────────────────────────

def norm(s) -> str:
    return str(s or "").strip().upper()

def read_individual_sheet(ws) -> List[dict]:
    raw = list(ws.iter_rows(values_only=True))
    if len(raw) < 4:
        return []

    col_map = {}

    def map_cell(n, i):
        if n == "QTD":                                          col_map[i] = "QTD"
        elif n in ("AREA", "AREA"):                             col_map[i] = "AREA"
        elif n in ("UNIDADE (FINAL)", "UNIDADE"):               col_map[i] = "UNIDADE"
        elif n == "POSTO/GRAD":                                 col_map[i] = "POSTO_GRAD"
        elif n == "QUADRO":                                     col_map[i] = "QUADRO"
        elif n == "NOME COMPLETO":                              col_map[i] = "NOME_COMPLETO"
        elif n == "RG":                                         col_map[i] = "RG"
        elif n in ("CAMISETA DE GV", "CAMISETA GV"):            col_map[i] = "CAMISETA_GV"
        elif n.startswith("CAMISA U"):                          col_map[i] = "CAMISA_UV"
        elif "SHORT JOHN" in n:                                 col_map[i] = "SHORT_JOHN"

    is_format_b = any(norm(c) == "QTD" for c in (raw[1] or []))
    if is_format_b:
        for i, c in enumerate(raw[1] or []): map_cell(norm(c), i)
        for i, c in enumerate(raw[2] or []): map_cell(norm(c), i)
    else:
        for i, c in enumerate(raw[2] or []): map_cell(norm(c), i)

    records = []
    for row in raw[3:]:
        if not row or all(str(v or "").strip() == "" for v in row):
            continue
        rec = {f: "" for f in ["QTD","AREA","UNIDADE","POSTO_GRAD","QUADRO","NOME_COMPLETO","RG","CAMISETA_GV","CAMISA_UV","SHORT_JOHN"]}
        for idx, val in enumerate(row):
            field = col_map.get(idx)
            if field:
                rec[field] = str(val or "").strip()
        rec["POSTO_GRAD"] = normalizar_posto(rec["POSTO_GRAD"]) or rec["POSTO_GRAD"]
        if rec["NOME_COMPLETO"]:
            records.append(rec)
    return records


# ── Endpoint: info da planilha consolidada ───────────────────────────────────

IGNORED_SHEETS = {"Resumo Geral", "Contagem", "Acrescentar1"}

@app.post("/info")
async def get_info(file: UploadFile = File(...)):
    data = await file.read()
    wb = load_workbook_safe(data, read_only=True, data_only=True)

    sheets = []
    for name in wb.sheetnames:
        if name in IGNORED_SHEETS:
            continue
        ws = wb[name]
        unidades = set()
        row_count = 0
        # Detect data start row (find QTD header then start after it)
        data_start = 4
        for ri, row in enumerate(ws.iter_rows(max_row=10, values_only=True), 1):
            if row and row[0] is not None and str(row[0]).strip().upper() == "QTD":
                data_start = ri + 1
                break
        for row in ws.iter_rows(min_row=data_start, values_only=True):
            if row and any(v is not None for v in row):
                row_count += 1
                if row[2]:
                    unidades.add(str(row[2]).strip())
        sheets.append({
            "name": name,
            "rows": row_count + 3,
            "unidades": sorted(unidades - {""})
        })
    wb.close()

    return {"sheets": sheets}


# ── Endpoint: mesclar planilha individual na consolidada ─────────────────────

@app.post("/merge")
async def merge(
    consolidada: UploadFile = File(...),
    individual:  UploadFile = File(...),
    aba_destino: str = Form(""),
    inserir_antes_de: str = Form(""),
):
    consolidada_bytes = await consolidada.read()
    individual_bytes  = await individual.read()

    wb_cons = load_workbook_safe(consolidada_bytes)
    wb_ind  = load_workbook_safe(individual_bytes)

    novos = []
    for sheet_name in wb_ind.sheetnames:
        ws_ind = wb_ind[sheet_name]
        novos.extend(read_individual_sheet(ws_ind))

    if not novos:
        raise HTTPException(400, "Nenhum militar encontrado na planilha individual.")

    aba_destino = aba_destino.strip()
    inserir_antes_de = inserir_antes_de.strip()
    if inserir_antes_de == "__fim__":
        inserir_antes_de = ""

    if aba_destino not in wb_cons.sheetnames:
        raise HTTPException(400, f"Aba '{aba_destino}' nao encontrada na planilha consolidada.")

    ws_dest = wb_cons[aba_destino]

    # Coleta areas reais presentes na aba destino para normalizacao
    areas_consolidada = []
    for row in ws_dest.iter_rows(min_row=4, values_only=True):
        if row and row[1]:
            val = str(row[1]).strip()
            if val and val not in areas_consolidada:
                areas_consolidada.append(val)

    # Normaliza area e posto de cada militar da planilha individual
    for mil in novos:
        mil["AREA"]       = normalizar_area(mil["AREA"], areas_consolidada)
        mil["UNIDADE"]    = normalizar_area(mil["UNIDADE"], areas_consolidada)
        mil["POSTO_GRAD"] = normalizar_posto(mil["POSTO_GRAD"]) or mil["POSTO_GRAD"]

    # Detect data start row dynamically (find QTD header)
    data_start_row = 4
    for ri in range(1, 10):
        v = ws_dest.cell(row=ri, column=1).value
        if v is not None and str(v).strip().upper() == "QTD":
            data_start_row = ri + 1
            break

    # Encontra linha de insercao
    insert_row = None
    if inserir_antes_de:
        for i in range(data_start_row, ws_dest.max_row + 1):
            val = ws_dest.cell(row=i, column=3).value
            if val and str(val).strip() == inserir_antes_de:
                insert_row = i
                break
    if insert_row is None:
        insert_row = ws_dest.max_row + 1

    n = len(novos)

    cores_originais = {}
    for r in range(data_start_row, ws_dest.max_row + 1):
        cores_originais[r] = get_color(ws_dest, r)

    ref = list(ws_dest[insert_row]) if insert_row <= ws_dest.max_row else []

    ws_dest.insert_rows(insert_row, amount=n)

    COR_AZUL   = "DAEEF3"
    COR_BRANCO = "FFFFFF"

    for i, mil in enumerate(novos):
        rn = insert_row + i
        cor = COR_AZUL if i % 2 == 0 else COR_BRANCO
        fill = make_fill(cor)
        values = [
            i + 1,
            mil["AREA"],
            mil["UNIDADE"],
            mil["POSTO_GRAD"],
            mil["QUADRO"],
            mil["NOME_COMPLETO"],
            mil["RG"],
            mil["CAMISETA_GV"],
            mil["CAMISA_UV"],
            mil["SHORT_JOHN"],
        ]
        for col_idx, val in enumerate(values, start=1):
            cell = ws_dest.cell(row=rn, column=col_idx, value=val)
            cell.fill = fill
            if col_idx <= len(ref) and ref[col_idx-1].font:
                cell.font = copy(ref[col_idx-1].font)
            if col_idx <= len(ref) and ref[col_idx-1].alignment:
                cell.alignment = copy(ref[col_idx-1].alignment)
            if col_idx <= len(ref) and ref[col_idx-1].border:
                cell.border = copy(ref[col_idx-1].border)

    for old_row, cor in cores_originais.items():
        new_row = old_row + n
        if new_row <= ws_dest.max_row:
            fill = make_fill(cor)
            for col in range(1, ws_dest.max_column + 1):
                ws_dest.cell(row=new_row, column=col).fill = fill

    # ── Atualiza aba Contagem ─────────────────────────────────────────────────
    if "Contagem" in wb_cons.sheetnames:
        ws_cont = wb_cons["Contagem"]

        unidades_novas = set(m["UNIDADE"] for m in novos if m["UNIDADE"])
        unidades_existentes = set()
        for i in range(1, ws_cont.max_row + 1):
            val = ws_cont.cell(row=i, column=1).value
            if val:
                unidades_existentes.add(str(val).strip())

        unidades_para_adicionar = sorted(unidades_novas - unidades_existentes)

        if unidades_para_adicionar:
            # Copia template do primeiro bloco (linhas 2-4) e adapta
            template_unidade = str(ws_cont.cell(row=2, column=1).value or "")
            template_rows = {}
            for offset in range(3):
                template_rows[offset] = {
                    col: ws_cont.cell(row=2 + offset, column=col).value
                    for col in range(1, ws_cont.max_column + 1)
                }

            insert_cont = ws_cont.max_row + 1
            if inserir_antes_de:
                for i in range(1, ws_cont.max_row + 1):
                    val = ws_cont.cell(row=i, column=1).value
                    if val and str(val).strip() == inserir_antes_de:
                        insert_cont = i
                        break

            rows_inseridos = len(unidades_para_adicionar) * 3
            ref_cont = list(ws_cont[max(1, insert_cont - 1)])
            ws_cont.insert_rows(insert_cont, amount=rows_inseridos)

            current_row = insert_cont
            for unidade in unidades_para_adicionar:
                for offset in range(3):
                    rn = current_row + offset
                    for col in range(1, ws_cont.max_column + 1):
                        tmpl_val = template_rows[offset].get(col)
                        cell = ws_cont.cell(row=rn, column=col)
                        if col == 1:
                            cell.value = unidade if offset == 0 else None
                        elif col == 2:
                            cell.value = tmpl_val
                        elif isinstance(tmpl_val, str) and tmpl_val.startswith("="):
                            new_f = tmpl_val
                            if template_unidade:
                                new_f = new_f.replace(
                                    f'"{template_unidade}"', f'"{unidade}"'
                                )
                            new_f = re.sub(r"'[^']*'!", f"'{aba_destino}'!", new_f)
                            new_f = re.sub(r"=SUM\(C\d+:I\d+\)", f"=SUM(C{rn}:I{rn})", new_f)
                            cell.value = new_f
                        if col <= len(ref_cont):
                            if ref_cont[col-1].font:      cell.font      = copy(ref_cont[col-1].font)
                            if ref_cont[col-1].alignment: cell.alignment = copy(ref_cont[col-1].alignment)
                            if ref_cont[col-1].border:    cell.border    = copy(ref_cont[col-1].border)
                current_row += 3

            # Corrige SUM deslocados pelo insert_rows
            for r in range(1, ws_cont.max_row + 1):
                cell_k = ws_cont.cell(row=r, column=11)
                val = str(cell_k.value or "")
                if not val.startswith("=SUM"):
                    continue
                m_sum = re.match(r"=SUM\(C(\d+):I(\d+)\)(.*)", val)
                if m_sum and int(m_sum.group(1)) != r:
                    cell_k.value = f"=SUM(C{r}:I{r}){m_sum.group(3)}"

    # ── Atualiza Resumo Geral ─────────────────────────────────────────────────
    if "Resumo Geral" in wb_cons.sheetnames and "Contagem" in wb_cons.sheetnames:
        ws_cont = wb_cons["Contagem"]
        ws_res  = wb_cons["Resumo Geral"]

        rows_cam, rows_uv, rows_sj = [], [], []
        for i in range(1, ws_cont.max_row + 1):
            mat = ws_cont.cell(row=i, column=2).value
            if mat == "Camiseta de GV": rows_cam.append(i)
            elif mat == "Camisa U.V":   rows_uv.append(i)
            elif mat == "Short John":   rows_sj.append(i)

        def build_sum(rows, cont_col):
            return "=" + "+".join(f"Contagem!{cont_col}{r}" for r in rows)

        cont_cols_cam = ["C","D","E","F","G","H","I","J"]
        for idx, cont_col in enumerate(cont_cols_cam):
            ws_res.cell(row=4, column=idx+2).value = build_sum(rows_cam, cont_col)
        ws_res.cell(row=4, column=10).value = build_sum(rows_cam, "K")

        cont_cols_uv = ["D","E","F","G","H"]
        for idx, cont_col in enumerate(cont_cols_uv):
            ws_res.cell(row=7, column=idx+2).value = build_sum(rows_uv, cont_col)
        ws_res.cell(row=7, column=7).value = build_sum(rows_uv, "K")

        for idx, cont_col in enumerate(cont_cols_uv):
            ws_res.cell(row=10, column=idx+2).value = build_sum(rows_sj, cont_col)
        ws_res.cell(row=10, column=7).value = build_sum(rows_sj, "K")

    # ── Retorna arquivo ───────────────────────────────────────────────────────
    output = io.BytesIO()
    wb_cons.save(output)
    output.seek(0)

    filename = consolidada.filename.replace(".xlsx", "_atualizado.xlsx")
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.get("/health")
def health():
    return {"status": "ok"}

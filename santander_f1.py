import os
import sys
import pdfplumber
import pandas as pd
from utils_santander import es_fecha, es_monto, monto_float, dist, unir_tokens_numericos, _formatear_excel

LEFT_BOUND, RIGHT_BOUND = 60.229, 102.0

def convertir_santander_f1(pdf_path: str, output_folder: str = None) -> str:
    if not output_folder:
        output_folder = sys.argv[1] if len(sys.argv) >= 2 else os.getcwd()
    with pdfplumber.open(pdf_path) as pdf:
        page0 = pdf.pages[1] if len(pdf.pages) > 1 else pdf.pages[0]
        empresa = (page0.within_bbox((43.5,71.729,369,76.63)).extract_text() or "").strip()
        no_cli  = (page0.within_bbox((479.98,77.378,528,82.279)).extract_text() or "").strip()
        hdrs = ["DEPOSITOS","RETIROS","SALDO"]
        col_pos, lines = {}, {}
        for w in page0.extract_words():
            lines.setdefault(int(w['top']), []).append(w)
        for _, ws in sorted(lines.items()):
            linea = " ".join(w['text'].upper() for w in ws)
            if all(h in linea for h in hdrs):
                for w in ws:
                    u = w['text'].upper()
                    if u in hdrs:
                        col_pos[u] = (w['x0']+w['x1'])/2
                break
        cols_sorted = sorted(col_pos.items(), key=lambda x: x[1])
        movs, mov = [], None
        read = stop = False
        periodo = rfc = ""
        skip = {"ESTADO DE CUENTA AL","PÁGINA"}
        stop_ph = {"INFORMACION FISCAL"}
        header_repeat = {"F E C H A FOLIO DESCRIPCION DEPOSITOS RETIROS SALDO","FECHA FOLIO DESCRIPCION DEPOSITOS RETIROS SALDO"}
        footer = {"BANCO SANTANDER (MEXICO)","INSTITUCION DE BANCA MULTIPLE,"}
        for idx, pg in enumerate(pdf.pages):
            if stop: break
            words = pg.extract_words()
            if idx>0: words = [w for w in words if w['top']>=36.25]
            lines = {}
            for w in words:
                lines.setdefault(int(w['top']), []).append(w)
            for _, ws in sorted(lines.items()):
                if stop: break
                line = " ".join(w['text'] for w in ws).strip(); u = line.upper()
                if "R.F.C." in u and not rfc:
                    rfc = " ".join(line.split()[line.split().index("R.F.C.")+1:]); continue
                if "PERIODO" in u and not periodo:
                    periodo = " ".join(line.split()[line.split().index("PERIODO")+1:]); continue
                if any(f in u for f in footer): break
                if any(sp in u for sp in stop_ph): stop = True; break
                if not read and any(es_fecha(t.upper()) for t in line.split()): read = True
                if not read: continue
                if any(s in u for s in (*skip,*header_repeat)): continue
                toks = u.split(); new_mov = toks and es_fecha(toks[0])
                if new_mov:
                    if mov: movs.append(mov)
                    mov = {"Fecha":toks[0],"Folio":None,"Descripción":"","Depositos":None,"Retiros":None,"Saldo":None,"_ft":[toks[0]]}
                elif not mov:
                    mov = {"Fecha":None,"Folio":None,"Descripción":"","Depositos":None,"Retiros":None,"Saldo":None}
                ws_proc = unir_tokens_numericos(ws) if new_mov else ws
                for w in ws_proc:
                    txt = w['text'].strip(); cx = (w['x0']+w['x1'])/2
                    if mov["Folio"] is None and LEFT_BOUND<=cx<=RIGHT_BOUND and txt.isdigit():
                        if txt not in mov["_ft"]: mov["Folio"]=txt; continue
                    if es_monto(txt):
                        val = monto_float(txt)
                        if cols_sorted:
                            col,_ = min(cols_sorted, key=lambda x:dist(x[1],cx))
                            mov[col.capitalize()] = val
                        else:
                            mov["Depositos"] = val
                    elif txt not in mov["_ft"]:
                        mov["Descripción"] += " "+txt
        if mov: movs.append(mov)
    df = pd.DataFrame(movs, columns=["Fecha","Folio","Descripción","Depositos","Retiros","Saldo"])
    excel_name = os.path.splitext(os.path.basename(pdf_path))[0]+".xlsx"
    excel_path = os.path.join(output_folder, excel_name)
    df.to_excel(excel_path, index=False)
    _formatear_excel(excel_path, empresa, no_cli, periodo, rfc)
    return excel_path
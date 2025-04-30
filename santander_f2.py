import os
import sys
import re
import pdfplumber
import pandas as pd
from utils_santander import es_fecha, es_monto, monto_float_2, _formatear_excel

DESCRIP_X0_MIN, DESCRIP_X0_MAX = 113.9, 350.0
TOTAL_H_MIN, TOTAL_H_MAX = 9.0, 11.0

DEBUG = True

def _agrupar_top(page): #igual
    d = {}
    for w in page.extract_words(): d.setdefault(int(w['top']), []).append(w)
    return d

def _detectar_columnas(pdf):
    busc={"DEPOSITO","RETIRO","SALDO"}
    for p in pdf.pages:
        for _, ws in sorted(_agrupar_top(p).items()):
            ln = " ".join(w['text'].upper() for w in ws)
            if all(x in ln for x in busc):
                return {w['text'].upper(): (w['x0']+w['x1'])/2 for w in ws if w['text'].upper() in busc}
    return {}

def convertir_santander_f2(pdf_path: str, output_folder: str = None) -> str:  #es procesar_archivos
    if not output_folder:
        output_folder = sys.argv[1] if len(sys.argv)>=2 else os.getcwd()
    movs = [] #era movimientos
    with pdfplumber.open(pdf_path) as pdf:

        if DEBUG and len(pdf.pages) > 10:
            print(f"\nDEBUG {os.path.basename(pdf_path)}  página 11 (índice 10)")
            for w in pdf.pages[10].extract_words():
                alto = w['bottom'] - w['top']
                print(f"Texto:{w['text']!r}  x0={w['x0']}  top={w['top']}  h={alto:.2f}")

        cols = _detectar_columnas(pdf)

        txt0 = pdf.pages[0].extract_text() or ""
        empresa=no_cli=periodo=rfc=""
        for ln in txt0.splitlines():
            u=ln.upper()
            if "CODIGO DE CLIENTE NO." in u and not no_cli: no_cli=u.split("CODIGO DE CLIENTE NO.")[1].strip()
            if "R.F.C." in u and not rfc: rfc=u.split("R.F.C.")[1].strip()
            if "PERIODO DEL" in u and not periodo:
                m=re.search(r"PERIODO DEL (.+)",ln,re.I)
                periodo=m.group(1).strip() if m else periodo
            if "DESARROLLADORA" in u and not empresa: empresa=ln.strip()
        start=False; mov=None; stop=False
        for pg in pdf.pages[1:]:
            if stop: break
            for _, ws in sorted(_agrupar_top(pg).items()):
                if stop: break
                ln=" ".join(w['text'].upper() for w in ws)
                if "FECHA FOLIO DESCRIPCION" in ln: continue
                if any(x in ln for x in ("ESTADO DE CUENTA","PÁGINA")): continue

                toks=ln.split()
                if toks and not start and es_fecha(toks[0][:11]): start=True
                if not start: continue
                
                if toks and es_fecha(toks[0][:11]):
                    if mov: movs.append(mov)
                    mov={"Fecha":toks[0][:11],"Folio":None,"Descripción":"","Depositos":None,"Retiros":None,"Saldo":None}
                if not mov: continue
                for w in ws:
                    txt=w['text'].strip(); alto=w['bottom']-w['top']
                    if txt.upper()=="TOTAL" and TOTAL_H_MIN<=alto<=TOTAL_H_MAX: stop=True; break
                    
                    cx=(w['x0']+w['x1'])/2

                    # ── Detección fecha+delimitador+folio (/, \ o |) ─────────
                    m = re.match(r'^(\d{2}-[A-Z]{3}-\d{4})[\/\\|](\d+)(.*)$', txt)
                   

                    if m:
                        fecha, folio, resto = m.group(1), m.group(2), m.group(3).strip()
                        mov["Fecha"] = fecha
                        mov["Folio"] = folio
                        if resto:
                            mov["Descripción"] += " " + resto
                        continue
                    
                    if re.match(r'^\d{2}-[A-Z]{3}-\d{4}', txt) and len(txt)>11:
                            fecha=txt[:11]
                            if es_fecha(fecha):
                                mov["Fecha"]=fecha
                                resto=txt[11:].strip()
                                m2=re.match(r'^(\d+)(.*)', resto)
                                if m2:
                                    mov["Folio"]=m2.group(1)
                                    resto=m2.group(2).strip()
                                if resto: mov["Descripción"]+=" "+resto
                            continue

                    # if mov["Folio"] is None and DESCRIP_X0_MIN<=w['x0']<=DESCRIP_X0_MAX and txt.isdigit(): mov["Folio"]=txt; continue
                    if mov["Folio"] is None and txt.isdigit():
                        mov["Folio"]=txt; continue
                    
                    if DESCRIP_X0_MIN <= w['x0'] <= DESCRIP_X0_MAX:
                        mov["Descripción"]+=" "+txt; continue

                    if es_monto(txt):
                        val=monto_float_2(txt)
                        if cols:
                            col=min(cols, key=lambda c:abs(cols[c]-cx))
                            if col=="DEPOSITO": mov["Depositos"]=val
                            elif col=="RETIRO": mov["Retiros"]=val
                            else: mov["Saldo"]=val
                        # else:
                        #     mov[col.capitalize()]=val
                        else:
                            mov["Depositos"]=val
                    else:
                        mov["Descripción"]+=" "+txt
                desc = mov["Descripción"].strip()
                desc = re.sub(r'^[\/\\|\s]+', '', desc)

                # 3) Quita repeticiones de la fecha al inicio
                fecha = mov.get("Fecha", "")
                desc = re.sub(rf'^(?:{re.escape(fecha)}\s*)+', '', desc)

                mov["Descripción"] = desc
                movs.append(mov)
            # if stop: break
        if mov and not stop: movs.append(mov)

    df = pd.DataFrame(movs, columns=["Fecha","Folio","Descripción","Depositos","Retiros","Saldo"])
    excel_name=os.path.splitext(os.path.basename(pdf_path))[0]+".xlsx"
    excel_path=os.path.join(output_folder, excel_name)
    df.to_excel(excel_path, index=False)
    _formatear_excel(excel_path, empresa, no_cli, periodo, rfc)
    return excel_path
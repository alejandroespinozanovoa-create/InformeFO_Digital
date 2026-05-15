import pandas as pd
import json
import numpy as np
import glob, os

# Find the Excel file
xlsx_files = glob.glob("*.xlsx") + glob.glob("*.XLSX")
if not xlsx_files:
    raise FileNotFoundError("No .xlsx file found in root")
xlsx_path = xlsx_files[0]
print(f"Processing: {xlsx_path}")

df = pd.read_excel(xlsx_path, sheet_name="BBDD", engine="openpyxl")
print(f"Loaded {len(df)} rows, {len(df.columns)} columns")

cols = [
    "DIA SEG","MES SEG","SEMANA","SUPERVISOR","LINEA",
    "Tramitador_Nom","Cedula_Tramitador",
    "ESTADO","TERMINADA","ANULADA","PENDIENTE","EMPAQUETADAS",
    "QUIEBRE","ATRIBUIBLE","REINGRESO",
    "ESTADO LEG","FECHA LEG","DIF DIAS","FINAL Vs LG","% PAGO",
    "ESTADO AGENDA","FECHA AGENDA","FRANJA AGENDA","DIA AGENDA","MES AGENDA",
    "Regional","Dpto_Nom","Mun_Nom","Depto",
    "Producto","TenenciaPlanta","Nom_Plan_Tarifario","ProductoPlanCD",
    "TarifaMXFinal","Tipo_Venta","TIPO_VENTA_DIG","CANAL_HOMO_DIG",
    "Dia_Reg","Mes_Reg","Anio_Reg",
    "DIAS MES PDC","DIA PDC",
]
existing = [c for c in cols if c in df.columns]
df2 = df[existing].copy()

for col in df2.columns:
    df2[col] = df2[col].apply(lambda x: x.strftime("%Y-%m-%d") if hasattr(x, "strftime") else x)
    df2[col] = df2[col].fillna("")

records = df2.to_dict(orient="records")

def clean(v):
    if isinstance(v, (np.integer,)): return int(v)
    if isinstance(v, (np.floating,)): return None if np.isnan(v) else float(v)
    if isinstance(v, np.bool_): return bool(v)
    return v

cleaned = [{k: clean(v) for k, v in row.items()} for row in records]

js_out = "const RAW_DATA = " + json.dumps(cleaned, ensure_ascii=False, separators=(",",":")) + ";"
with open("data.js", "w", encoding="utf-8") as f:
    f.write(js_out)
print(f"Written data.js: {len(cleaned)} records, {len(js_out):,} bytes")

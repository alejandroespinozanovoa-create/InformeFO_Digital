import pandas as pd
import json
import numpy as np
import glob

xlsx_files = glob.glob("*.xlsx") + glob.glob("*.XLSX")
if not xlsx_files:
    raise FileNotFoundError("No .xlsx found in root")
xlsx_path = xlsx_files[0]
print(f"Processing: {xlsx_path}")

df     = pd.read_excel(xlsx_path, sheet_name="BBDD",     engine="openpyxl")
df_can = pd.read_excel(xlsx_path, sheet_name="BBDD CAN", engine="openpyxl")
df_leg = pd.read_excel(xlsx_path, sheet_name="BBDD LEG", engine="openpyxl")
df_pla = pd.read_excel(xlsx_path, sheet_name="PLANTA",   engine="openpyxl")
df_bck = pd.read_excel(xlsx_path, sheet_name="BACK",     engine="openpyxl")
print(f"BBDD: {len(df)} rows x {len(df.columns)} cols")

def clean(v):
    if isinstance(v, (np.integer,)):  return int(v)
    if isinstance(v, (np.floating,)): return None if np.isnan(v) else round(float(v),4)
    if isinstance(v, np.bool_):       return bool(v)
    if hasattr(v, "strftime"):        return v.strftime("%Y-%m-%d")
    if hasattr(v, "hour"):            return None
    return v

def to_recs(frame, cols=None):
    if cols:
        frame = frame[[c for c in cols if c in frame.columns]].copy()
    else:
        frame = frame.copy()
    for col in frame.columns:
        frame[col] = frame[col].apply(lambda x: x.strftime("%Y-%m-%d") if hasattr(x,"strftime") else (None if hasattr(x,"hour") else x))
        frame[col] = frame[col].fillna("")
    return [{k: clean(v) for k,v in r.items()} for r in frame.to_dict(orient="records")]

BBDD_COLS = [
    "DIA SEG","MES SEG","SEMANA","QUIEBRE","ATRIBUIBLE","REINGRESO",
    "SUPERVISOR","LINEA","ESTADO","TERMINADA","ANULADA","PENDIENTE",
    "ESTADO AGENDA","FECHA AGENDA","FRANJA AGENDA","DIA AGENDA","MES AGENDA",
    "MOTIVO DE SUSPENSION","CEDULA CLIENTE","TELEFONO CLIENTE",
    "ESTADO LEG","DIF DIAS","FECHA LEG","FINAL Vs LG","% PAGO",
    "DIA PDC","DIAS MES PDC","EMPAQUETADAS",
    "META","BA","DUO","TRIO",
    "Tramitador_Nom","Cedula_Tramitador","FechaRegistroPeticion",
    "Regional","Dpto_Nom","Mun_Nom","Localidad","Depto",
    "Producto","TenenciaPlanta","Nom_Plan_Tarifario",
    "TarifaMXFinal","Tipo_Venta","Convergente","Tipo_Convergente","Tipo_Cliente",
    "SubSegmentoCuentaDesc","Estrato","Zona",
    "UNIDAD_VENTA_DIG","ATR_VENTA_SOURCE","SubCanal",
    "ESTADO_VARIABLE_DIG","MOTIVO_QUIEBRE_DIG",
    "Altas_Con_Reingreso","Dia_Reg","Mes_Reg","Anio_Reg",
    "Nom_Ult_Quiebre_Atmp","TipoQuiebreDesc","SubCanal_Hom",
    "Nom_Punto_Venta","Nom_Barrio_PC","Oferta_Empaquetada",
]

bbdd_recs = to_recs(df, BBDD_COLS)
can_recs  = to_recs(df_can, ["ID","QUIEBRE","ATRIBUIBLE","Nom_Ult_Quiebre_Atmp","TipoQuiebreDesc","Remark"])
leg_recs  = to_recs(df_leg, ["NUM_PETICION","FEC_ALTA","FEC_REGISTRO","NOM_VENDEDOR","NUM_IDENT_VENDEDOR",
                              "MUNICIPIO","DEPARTAMENTO","REGIONAL","COMERCIALIZADOR","NOMBRE_GRUPO",
                              "MULTIPRODUCTO","Nom_Plan_Tarifario","DIF_DIAS","FEC_LEGAL",
                              "ESTADO_LEG","SITUACION_ORDEN","ESTADO_VENDEDOR","SubSegmentoCuentaDesc","RANGO","ST"])
pla_recs  = to_recs(df_pla, ["FECHA","CEDULA","NOMBRE","SUPERVISOR","LINEA"])
bck_recs  = to_recs(df_bck[["QUIEBRE","ATRIBUIBLE","PROCESO"]].dropna(subset=["QUIEBRE"]), ["QUIEBRE","ATRIBUIBLE","PROCESO"])

config = {
    "DIA_PDC":      int(df["DIA PDC"].iloc[0])      if "DIA PDC"      in df.columns else 12,
    "DIAS_MES_PDC": int(df["DIAS MES PDC"].iloc[0]) if "DIAS MES PDC" in df.columns else 24,
    "META": {"GEN":1147,"FMC":491,"TVT":48},
}

payload = json.dumps({"BBDD":bbdd_recs,"CAN":can_recs,"LEG":leg_recs,
                      "PLANTA":pla_recs,"QUIEBRE_CATALOG":bck_recs,"CONFIG":config},
                     ensure_ascii=False, separators=(",",":"))

with open("data.js","w",encoding="utf-8") as f:
    f.write("const DB="+payload+";")

print(f"data.js: {len(payload):,} bytes | BBDD:{len(bbdd_recs)} CAN:{len(can_recs)} LEG:{len(leg_recs)} PLANTA:{len(pla_recs)}")

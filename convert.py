import pandas as pd
import json
import numpy as np
import glob

# ─── Find Excel ─────────────────────────────────────────────────────────
xlsx_files = glob.glob("*.xlsx") + glob.glob("*.XLSX")
if not xlsx_files:
    raise FileNotFoundError("No .xlsx found in root")
xlsx_path = xlsx_files[0]
print(f"Processing: {xlsx_path}")

df = pd.read_excel(xlsx_path, sheet_name="BBDD", engine="openpyxl")
print(f"BBDD: {len(df)} rows x {len(df.columns)} cols")

# ─── Clean helper ────────────────────────────────────────────────────────
def clean(v):
    if isinstance(v, (np.integer,)):  return int(v)
    if isinstance(v, (np.floating,)): return None if np.isnan(v) else round(float(v), 4)
    if isinstance(v, np.bool_):       return bool(v)
    if hasattr(v, "strftime"):        return v.strftime("%Y-%m-%d")
    if hasattr(v, "hour"):            return None  # datetime.time objects
    return v

def to_recs(frame, cols):
    use = [c for c in cols if c in frame.columns]
    frame = frame[use].copy()
    for col in frame.columns:
        frame[col] = frame[col].apply(
            lambda x: x.strftime("%Y-%m-%d") if hasattr(x, "strftime")
                      else (None if hasattr(x, "hour") else x)
        )
        frame[col] = frame[col].fillna("")
    return [{k: clean(v) for k, v in r.items()} for r in frame.to_dict(orient="records")]

# ─── All relevant columns from BBDD ─────────────────────────────────────
# First 37 are custom/formula columns added by the user - highest priority
BBDD_COLS = [
    # ── Custom cols (user-added formulas, cols 0-36) ────────────────────
    "DIA SEG", "MES SEG", "SEMANA",
    "QUIEBRE", "ATRIBUIBLE", "REINGRESO",
    "SUPERVISOR", "LINEA", "ESTADO",
    "TERMINADA", "ANULADA", "PENDIENTE",
    "AGENDADAS", "PENDIENTES", "SUSPENDIDAS",
    "ESTADO AGENDA", "FECHA AGENDA", "FRANJA AGENDA",
    "DIA AGENDA", "MES AGENDA", "MOTIVO DE SUSPENSION",
    "CEDULA CLIENTE", "TELEFONO CLIENTE",
    "LEGALIZADA", "PEND LEGALIZAR",
    "ESTADO LEG", "DIF DIAS", "FECHA LEG", "FINAL Vs LG", "% PAGO",
    "DIA PDC", "DIAS MES PDC", "EMPAQUETADAS",
    "META", "BA", "DUO", "TRIO",
    # ── Original source columns ─────────────────────────────────────────
    "Tramitador_Nom", "Cedula_Tramitador",
    "FechaRegistroPeticion",
    "Regional", "Dpto_Nom", "Mun_Nom", "Localidad", "Depto",
    "Producto", "TenenciaPlanta", "Nom_Plan_Tarifario",
    "TarifaMXFinal", "Tipo_Venta", "Convergente", "Tipo_Convergente", "Tipo_Cliente",
    "SubSegmentoCuentaDesc", "Estrato", "Zona",
    "UNIDAD_VENTA_DIG", "ATR_VENTA_SOURCE", "SubCanal",
    "ESTADO_VARIABLE_DIG", "MOTIVO_QUIEBRE_DIG",
    "Altas_Con_Reingreso",
    "Dia_Reg", "Mes_Reg", "Anio_Reg",
    "Nom_Ult_Quiebre_Atmp", "TipoQuiebreDesc",
    "SubCanal_Hom", "Nom_Punto_Venta", "Nom_Barrio_PC",
    "Oferta_Empaquetada",
    "REMARK",
    # Quiebre catalog cols (embedded in BBDD)
    "Nom_PS", "ProductoPlanCD",
]

bbdd_recs = to_recs(df, BBDD_COLS)

# ─── Build quiebre catalog from BBDD itself ───────────────────────────
# Map motivo -> atribuible from the data
quiebre_cat = []
if "QUIEBRE" in df.columns and "ATRIBUIBLE" in df.columns:
    q_df = df[["QUIEBRE","ATRIBUIBLE"]].dropna(subset=["QUIEBRE"])
    q_df = q_df[q_df["QUIEBRE"].astype(str).str.strip() != ""]
    seen = set()
    for _, row in q_df.iterrows():
        k = str(row["QUIEBRE"]).strip().upper()
        if k not in seen:
            seen.add(k)
            quiebre_cat.append({
                "QUIEBRE": str(row["QUIEBRE"]).strip(),
                "ATRIBUIBLE": str(row["ATRIBUIBLE"]).strip() if pd.notna(row["ATRIBUIBLE"]) else "",
                "PROCESO": ""
            })

# ─── Config ──────────────────────────────────────────────────────────────
config = {
    "DIA_PDC":      int(df["DIA PDC"].iloc[0])      if "DIA PDC"      in df.columns else 1,
    "DIAS_MES_PDC": int(df["DIAS MES PDC"].iloc[0]) if "DIAS MES PDC" in df.columns else 24,
    "META":         {"GEN": 1147, "FMC": 491, "TVT": 48},
}

# ─── Write data.js ────────────────────────────────────────────────────────
payload = json.dumps(
    {
        "BBDD":            bbdd_recs,
        "CAN":             [],          # empty — data comes from BBDD.QUIEBRE
        "LEG":             [],          # empty — data comes from BBDD.ESTADO LEG
        "PLANTA":          [],          # empty — no PLANTA sheet
        "QUIEBRE_CATALOG": quiebre_cat,
        "CONFIG":          config,
    },
    ensure_ascii=False,
    separators=(",", ":")
)

with open("data.js", "w", encoding="utf-8") as f:
    f.write("const DB=" + payload + ";")

print(f"data.js written: {len(payload):,} bytes | {len(bbdd_recs)} records | {len(quiebre_cat)} quiebre types")

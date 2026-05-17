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

# ─── Compute DIA SEG / MES SEG if not present ───────────────────────────
# Some Excel versions don't include these formula columns; derive them.
if "DIA SEG" not in df.columns:
    if "Dia_Seg" in df.columns:
        df["DIA SEG"] = pd.to_numeric(df["Dia_Seg"], errors="coerce")
    elif "FechaRegistroPeticion" in df.columns:
        df["DIA SEG"] = pd.to_datetime(df["FechaRegistroPeticion"], errors="coerce").dt.day
    print("  ↳ DIA SEG computed from source columns")

if "MES SEG" not in df.columns:
    if "Mes_Seg" in df.columns:
        df["MES SEG"] = pd.to_numeric(df["Mes_Seg"], errors="coerce")
    elif "FechaRegistroPeticion" in df.columns:
        df["MES SEG"] = pd.to_datetime(df["FechaRegistroPeticion"], errors="coerce").dt.month
    print("  ↳ MES SEG computed from source columns")

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
BBDD_COLS = [
    # ── Custom / formula cols ───────────────────────────────────────────
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
    # ── Source columns ──────────────────────────────────────────────────
    "Tramitador_Nom", "Cedula_Tramitador",
    "FechaRegistroPeticion",
    "Regional", "Dpto_Nom", "Mun_Nom", "Localidad", "Depto",
    "Producto", "TenenciaPlanta", "Nom_Plan_Tarifario",
    "TarifaMXFinal", "Tipo_Venta", "Convergente", "Tipo_Convergente", "Tipo_Cliente",
    "SubSegmentoCuentaDesc", "Estrato", "Zona",
    "UNIDAD_VENTA_DIG", "ATR_VENTA_SOURCE", "ATR_VENTA_CAMPANA", "ATR_VENTA_MEDIUM", "SubCanal",
    "ESTADO_VARIABLE_DIG", "MOTIVO_QUIEBRE_DIG",
    "Altas_Con_Reingreso",
    "Dia_Reg", "Mes_Reg", "Anio_Reg",
    "Nom_Ult_Quiebre_Atmp", "TipoQuiebreDesc",
    "SubCanal_Hom", "Nom_Punto_Venta", "Nom_Barrio_PC",
    "Oferta_Empaquetada",
    "REMARK",
    "Nom_PS", "ProductoPlanCD",
]

bbdd_recs = to_recs(df, BBDD_COLS)
print(f"  ↳ BBDD: {len(bbdd_recs)} registros")

# ─── Build QUIEBRE_CATALOG ────────────────────────────────────────────
quiebre_cat = []
if "QUIEBRE" in df.columns and "ATRIBUIBLE" in df.columns:
    q_df = df[["QUIEBRE", "ATRIBUIBLE"]].dropna(subset=["QUIEBRE"])
    q_df = q_df[q_df["QUIEBRE"].astype(str).str.strip() != ""]
    seen = set()
    for _, row in q_df.iterrows():
        k = str(row["QUIEBRE"]).strip().upper()
        if k not in seen:
            seen.add(k)
            quiebre_cat.append({
                "QUIEBRE":    str(row["QUIEBRE"]).strip(),
                "ATRIBUIBLE": str(row["ATRIBUIBLE"]).strip() if pd.notna(row["ATRIBUIBLE"]) else "",
                "PROCESO":    ""
            })
print(f"  ↳ QUIEBRE_CATALOG: {len(quiebre_cat)} tipos")

# ─── Build CAN (cancellations / quiebres) from BBDD ─────────────────
# The index.html uses DB.CAN for the quiebre analysis tab.
# CAN = BBDD rows where QUIEBRE is filled OR ANULADA = 1.
CAN_COLS = [
    "QUIEBRE", "ATRIBUIBLE",
    "SUPERVISOR", "LINEA",
    "Tramitador_Nom", "Cedula_Tramitador",
    "Nom_Ult_Quiebre_Atmp", "TipoQuiebreDesc",
    "SEMANA", "DIA SEG", "MES SEG",
    "Regional", "Dpto_Nom", "Mun_Nom",
    "Producto", "ESTADO",
    "TERMINADA", "ANULADA", "PENDIENTE",
    "FechaRegistroPeticion", "Dia_Reg", "Mes_Reg", "Anio_Reg",
    "MOTIVO_QUIEBRE_DIG",
]

can_mask = pd.Series([False] * len(df))
if "QUIEBRE" in df.columns:
    can_mask |= df["QUIEBRE"].astype(str).str.strip().ne("")
if "ANULADA" in df.columns:
    can_mask |= pd.to_numeric(df["ANULADA"], errors="coerce").fillna(0).astype(int).eq(1)

can_df = df[can_mask].copy()
can_recs = to_recs(can_df, CAN_COLS)
print(f"  ↳ CAN: {len(can_recs)} registros de quiebre/anulación")

# ─── Build LEG (legalizations) from BBDD ────────────────────────────
# index.html uses DB.LEG with fields: FEC_ALTA (date string), SITUACION_ORDEN
# SITUACION_ORDEN values expected: 'MANUAL', 'AUTOMATICA', 'PENDIENTE LEG'
# Mapped from BBDD column 'ESTADO LEG':
#   'legalizada grabacion' → 'MANUAL'
#   'legalizada aut'       → 'AUTOMATICA'
#   'pdte legalizar'       → 'PENDIENTE LEG'
#   'no cruza'             → 'NO CRUZA'

def map_situacion(estado_leg):
    s = str(estado_leg).strip().lower()
    if "grabacion" in s or (s.startswith("legalizada") and "aut" not in s):
        return "MANUAL"
    if "aut" in s and "legalizada" in s:
        return "AUTOMATICA"
    if "pdte" in s or "pendiente" in s:
        return "PENDIENTE LEG"
    if "no cruza" in s:
        return "NO CRUZA"
    return "PENDIENTE LEG"

leg_recs = []
for r in bbdd_recs:
    estado_leg = str(r.get("ESTADO LEG") or "").strip()
    fecha_leg  = r.get("FECHA LEG") or ""
    if not estado_leg and not fecha_leg:
        continue  # skip rows with no legalization data

    leg_recs.append({
        "FEC_ALTA":       fecha_leg,               # date of legalization
        "SITUACION_ORDEN": map_situacion(estado_leg),
        "ESTADO LEG":     estado_leg,              # original value kept for reference
        "Tramitador_Nom": r.get("Tramitador_Nom") or "",
        "SUPERVISOR":     r.get("SUPERVISOR") or "",
        "LINEA":          r.get("LINEA") or "",
        "% PAGO":         r.get("% PAGO") or "",
        "DIF DIAS":       r.get("DIF DIAS") or "",
        "FINAL Vs LG":    r.get("FINAL Vs LG") or "",
        "SEMANA":         r.get("SEMANA") or "",
        "DIA SEG":        r.get("DIA SEG") or "",
        "MES SEG":        r.get("MES SEG") or "",
        "Regional":       r.get("Regional") or "",
    })

print(f"  ↳ LEG: {len(leg_recs)} registros de legalización")

# ─── Build PLANTA (agent roster) from BBDD unique agents ────────────
# index.html uses DB.PLANTA with: NOMBRE, SUPERVISOR, LINEA
# Build from unique Tramitador_Nom + SUPERVISOR + LINEA combinations
planta_recs = []
seen_agents = set()
for r in bbdd_recs:
    nombre = str(r.get("Tramitador_Nom") or "").strip()
    if nombre and nombre not in seen_agents:
        seen_agents.add(nombre)
        planta_recs.append({
            "NOMBRE":     nombre,
            "SUPERVISOR": r.get("SUPERVISOR") or "",
            "LINEA":      r.get("LINEA") or "",
            "CEDULA":     r.get("Cedula_Tramitador") or "",
        })

print(f"  ↳ PLANTA: {len(planta_recs)} agentes únicos")

# ─── Config ──────────────────────────────────────────────────────────────
dia_pdc = 1
dias_mes_pdc = 24

if "DIA PDC" in df.columns:
    val = pd.to_numeric(df["DIA PDC"], errors="coerce").dropna()
    if not val.empty:
        dia_pdc = int(val.iloc[0])

if "DIAS MES PDC" in df.columns:
    val = pd.to_numeric(df["DIAS MES PDC"], errors="coerce").dropna()
    if not val.empty:
        dias_mes_pdc = int(val.iloc[0])

# ─── Extract META per supervisor from BBDD ──────────────────────────────
meta_sup_dict = {}
if "META" in df.columns and "SUPERVISOR" in df.columns:
    grp = df[df["SUPERVISOR"].astype(str).str.strip() != ""].groupby("SUPERVISOR")["META"]
    for sup, vals in grp:
        v = pd.to_numeric(vals, errors="coerce").dropna()
        if not v.empty:
            meta_val = int(v.iloc[0])
            if meta_val > 0:
                meta_sup_dict[str(sup).strip()] = meta_val
meta_gen = sum(meta_sup_dict.values())

config = {
    "DIA_PDC":      dia_pdc,
    "DIAS_MES_PDC": dias_mes_pdc,
    "META_GEN":     meta_gen,
    "META_SUP":     meta_sup_dict,
}

print(f"  ↳ CONFIG: DIA_PDC={dia_pdc}, DIAS_MES_PDC={dias_mes_pdc}, META_GEN={meta_gen}")
print(f"  ↳ META_SUP: {meta_sup_dict}")

# ─── Write data.js ───────────────────────────────────────────────────────
payload = json.dumps(
    {
        "BBDD":            bbdd_recs,
        "CAN":             can_recs,
        "LEG":             leg_recs,
        "PLANTA":          planta_recs,
        "QUIEBRE_CATALOG": quiebre_cat,
        "CONFIG":          config,
    },
    ensure_ascii=False,
    separators=(",", ":")
)

with open("data.js", "w", encoding="utf-8") as f:
    f.write("const DB=" + payload + ";")

print(f"\n✅ data.js escrito exitosamente:")
print(f"   Tamaño  : {len(payload):,} bytes")
print(f"   BBDD    : {len(bbdd_recs):,} registros")
print(f"   CAN     : {len(can_recs):,} quiebres/anulaciones")
print(f"   LEG     : {len(leg_recs):,} legalizaciones")
print(f"   PLANTA  : {len(planta_recs):,} agentes")
print(f"   QUIEBRE : {len(quiebre_cat):,} tipos de quiebre")

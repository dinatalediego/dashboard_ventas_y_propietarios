from pathlib import Path
import pandas as pd
import re


# =========================
# CONFIGURACIÓN
# =========================

INPUT_FILE = Path("data/leads_crm.xlsx")
OUTPUT_FILE = Path("outputs/bbdd_negativa_exclusion.xlsx")

# Ajusta columnas reales
COL_NOMBRE = "nombre"
COL_DOCUMENTO = "documento"
COL_CORREO = "correo"
COL_TELEFONO = "telefono"
COL_PROYECTO = "proyecto"
COL_ESTADO = "estado_lead"
COL_OBSERVACION = "observacion"
COL_FECHA = "fecha_registro"


# =========================
# FUNCIONES
# =========================

def load_file(path):
    suffix = path.suffix.lower()

    if suffix in [".xlsx", ".xls"]:
        return pd.read_excel(path)

    if suffix == ".csv":
        return pd.read_csv(path, encoding="utf-8-sig")

    raise ValueError("Formato no soportado. Usa Excel o CSV.")


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_text(value):
    value = clean_text(value).upper()
    replacements = {
        "Á": "A",
        "É": "E",
        "Í": "I",
        "Ó": "O",
        "Ú": "U",
        "Ü": "U",
    }
    for old, new in replacements.items():
        value = value.replace(old, new)
    return value


def normalize_phone(value):
    if pd.isna(value):
        return ""

    phone = re.sub(r"\D", "", str(value))

    if phone.startswith("51") and len(phone) == 11:
        phone = phone[2:]

    return phone


def normalize_email(value):
    return clean_text(value).lower()


def safe_get(row, col):
    if col in row.index:
        return row.get(col, "")
    return ""


def detect_negative_reason(row):
    estado = normalize_text(safe_get(row, COL_ESTADO))
    observacion = normalize_text(safe_get(row, COL_OBSERVACION))

    texto = f"{estado} {observacion}"

    if "NO PERFIL" in texto or "SIN PERFIL" in texto or "NO CALZA PERFIL" in texto:
        return "NO perfil"

    if (
        "NO ALCANZA CUOTA" in texto
        or "NO CALIFICA" in texto
        or "NO CALIFICO" in texto
        or "NO APTO" in texto
        or "NO SUJETO" in texto
        or "NO CUMPLE CUOTA" in texto
    ):
        return "No alcanza cuota"

    if (
        "NO RESPONDIO" in texto
        or "NO RESPONDE" in texto
        or "NO CONTACTADO" in texto
        or "NO CONTESTA" in texto
        or "NO CONTESTO" in texto
        or "SIN RESPUESTA" in texto
        or "NO UBICADO" in texto
    ):
        return "No respondió"

    return None


def add_missing_columns(df, cols):
    for col in cols:
        if col not in df.columns:
            df[col] = ""


# =========================
# PROCESO
# =========================

def main():
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"No encontré el archivo {INPUT_FILE}. "
            "Coloca tu archivo en data/ y renómbralo como leads_crm.xlsx"
        )

    df = load_file(INPUT_FILE)

    add_missing_columns(
        df,
        [
            COL_NOMBRE,
            COL_DOCUMENTO,
            COL_CORREO,
            COL_TELEFONO,
            COL_PROYECTO,
            COL_ESTADO,
            COL_OBSERVACION,
            COL_FECHA,
        ]
    )

    df["telefono_exclusion"] = df[COL_TELEFONO].apply(normalize_phone)
    df["correo_exclusion"] = df[COL_CORREO].apply(normalize_email)
    df["documento_exclusion"] = df[COL_DOCUMENTO].apply(clean_text)

    df["motivo_exclusion"] = df.apply(detect_negative_reason, axis=1)

    negativos = df[df["motivo_exclusion"].notna()].copy()

    cols_finales = [
        COL_NOMBRE,
        "documento_exclusion",
        "correo_exclusion",
        "telefono_exclusion",
        COL_PROYECTO,
        COL_FECHA,
        "motivo_exclusion",
        COL_ESTADO,
        COL_OBSERVACION,
    ]

    cols_existentes = [c for c in cols_finales if c in negativos.columns]

    negativos_final = negativos[cols_existentes].copy()

    subset_dedup = [
        c for c in [
            "documento_exclusion",
            "correo_exclusion",
            "telefono_exclusion"
        ]
        if c in negativos_final.columns
    ]

    if subset_dedup:
        negativos_final = negativos_final.drop_duplicates(
            subset=subset_dedup,
            keep="last"
        )

    resumen_motivo = (
        negativos_final
        .groupby("motivo_exclusion", dropna=False)
        .size()
        .reset_index(name="cantidad")
        .sort_values("cantidad", ascending=False)
    )

    resumen_proyecto = (
        negativos_final
        .groupby([COL_PROYECTO, "motivo_exclusion"], dropna=False)
        .size()
        .reset_index(name="cantidad")
        .sort_values([COL_PROYECTO, "cantidad"], ascending=[True, False])
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        negativos_final.to_excel(writer, index=False, sheet_name="bbdd_negativa")
        resumen_motivo.to_excel(writer, index=False, sheet_name="resumen_motivo")
        resumen_proyecto.to_excel(writer, index=False, sheet_name="resumen_proyecto")

    print(f"OK - Archivo generado: {OUTPUT_FILE}")
    print(f"Contactos negativos encontrados: {len(negativos_final)}")


if __name__ == "__main__":
    main()

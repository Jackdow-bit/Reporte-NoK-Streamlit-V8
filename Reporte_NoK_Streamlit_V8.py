# -*- coding: utf-8 -*-
"""
Reporte NoK - Streamlit V8

Procesamiento masivo de reportes PDF:
- PDFs individuales: múltiples archivos, incluso miles.
- Uno o varios ZIP.
- Mezcla de PDFs + ZIP en una misma carga.
- ZIP anidados hasta MAX_ZIP_DEPTH.
- Detección de duplicados por SHA-256.
- Procesamiento paralelo con ThreadPoolExecutor.
- Reporte Excel con hojas de resultados y análisis.
- Visualización interactiva en Streamlit.

Basado en la lógica del notebook proporcionado por el usuario.
"""

from __future__ import annotations

import hashlib
import io
import os
import re
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from typing import Iterable

import numpy as np
import pandas as pd
import pdfplumber
import streamlit as st

# =========================
# CONFIGURACIÓN
# =========================
APP_VERSION = "V8.0"
MAX_ZIP_DEPTH = 8
DEFAULT_WORKERS = min(12, max(4, (os.cpu_count() or 4) * 2))
MAX_EXCEL_ROWS = 1_048_575  # Excel: 1,048,576 filas incluyendo encabezado

# Umbral solicitado: Levas (5) 301-400 UPR = >= 0.0001 => NOK
CHATTER_LEVAS_THRESHOLD = 0.0001
# Lógica de Apoyos se mantiene separada; actualmente las características
# de Chatter Apoyos no usan (5)/(9) 301-400 UPR.
CHATTER_APOYOS_THRESHOLD = 0.00008

CHARACTERISTICS_FOR_COLUMNS = [
    "Diametro", "Roundness", "Runout", "Concentricity",
    "Parallelism", "Taper", "Cylindricity"
]

APOYO_IDENTIFIERS = ["Aux:G", "1:A L", "1:A R", "2:C", "3:D", "4:E", "5:B"]

APOYO_RENAME_MAPPING = {
    "1:": "1:A L",
    "2:": "1:A R",
    "3:": "2:C",
    "4:": "3:D",
    "5:": "4:E",
    "6:": "5:B",
}

CHATTER_LEVAS_CHARACTERISTICS = [
    "(1) 40- 80 UPR", "(2) 81-140 UPR", "(3) 141-190 UPR",
    "(4) 191-300 UPR", "(5) 301-400 UPR"
]

CHATTER_APOYOS_CHARACTERISTICS = [
    "(1) 5 - 8 UPR", "(2) 9 - 15 UPR", "(3) 16 - 23 UPR",
    "(4) 24 - 28 UPR", "(5) 29- 45 UPR", "(6) 46-70 UPR",
    "(7) 71-140 UPR", "(8) 141-215 UPR", "(9) 216-270 UPR",
    "(10) 271-300 UPR"
]

LOBES_HEADERS = [
    "AngleErr", "BC-Rad'sErr", "BC-Runout", "BC-Vel./10°",
    "Ramp-MaxLift", "Nose-MaxLift", "Ramp+9°Vel/1°",
    "Nose-Vel./1°", "Taper", "Center-Dev"
]

CHATTER_FORD_THRESHOLDS = {
    "(1) 29 - 45 UPR": 0.00033,
    "(2) 46 -70 UPR": 0.00033,
    "(3) 71 - 140 UPR": 0.00016,
    "(4) 141 - 215 UPR": 0.00013,
    "(5) 216 - 270 UPR": 0.00008,
}

CHATTER_FORD_ORIGINAL = [
    "(5) 29- 45 UPR", "(6) 46-70 UPR", "(7) 71-140 UPR",
    "(8) 141-215 UPR", "(9) 216-270 UPR"
]
CHATTER_FORD_RENAME = {
    "(5) 29- 45 UPR": "(1) 29 - 45 UPR",
    "(6) 46-70 UPR": "(2) 46 -70 UPR",
    "(7) 71-140 UPR": "(3) 71 - 140 UPR",
    "(8) 141-215 UPR": "(4) 141 - 215 UPR",
    "(9) 216-270 UPR": "(5) 216 - 270 UPR",
}

OUTPUT_COLUMNS = [
    "Nombre del archivo", "Pieza", "Leva", "Apoyo", "Caracteristica",
    "Medicion", "Area", "Resultado"
]


@dataclass(frozen=True)
class PDFJob:
    name: str
    data: bytes
    sha256: str


# =========================
# REGLAS DE RESULTADO
# =========================
def _to_float(value) -> float | None:
    try:
        return float(str(value).replace(",", ".").strip())
    except (ValueError, TypeError):
        return None


def get_resultado(value_str, characteristic_name=None, area=None):
    """Regla principal del reporte.

    '#' siempre es NOK.
    En Chatter Levas, (5) 301-400 UPR es NOK cuando valor >= 0.0001.
    """
    text_value = str(value_str)
    if "#" in text_value:
        return "Nok"

    if characteristic_name == "(5) 301-400 UPR" and area in ("Base Circle", "Lift Area"):
        numeric_val = _to_float(text_value)
        if numeric_val is not None and numeric_val >= CHATTER_LEVAS_THRESHOLD:
            return "Nok"

    # Se conserva la regla para (9) por compatibilidad con versiones anteriores.
    if characteristic_name == "(9) 301-400 UPR":
        numeric_val = _to_float(text_value)
        if numeric_val is not None and numeric_val >= CHATTER_APOYOS_THRESHOLD:
            return "Nok"

    return "Ok"


def get_resultado_chatter_journal_ford(medicion_val, characteristic_name):
    if "#" in str(medicion_val):
        return "Nok"
    threshold = CHATTER_FORD_THRESHOLDS.get(characteristic_name)
    if threshold is None:
        return "Ok"
    numeric_val = _to_float(medicion_val)
    if numeric_val is None:
        return "Nok"
    return "Nok" if numeric_val > threshold else "Ok"


# =========================
# UTILIDADES DE PDF / ZIP
# =========================
def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def iter_uploaded_sources(uploaded_files: Iterable, max_depth: int = MAX_ZIP_DEPTH):
    """Devuelve PDFs encontrados en uploads y ZIPs, sin extraer a disco."""
    stats = {"pdf_candidates": 0, "zip_files": 0, "invalid": 0, "nested_zips": 0}

    def walk(name: str, data: bytes, depth: int):
        lower = name.lower()
        if lower.endswith(".pdf"):
            stats["pdf_candidates"] += 1
            yield name, data
            return

        if lower.endswith(".zip"):
            stats["zip_files"] += 1
            if depth >= max_depth:
                return
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as zf:
                    for info in zf.infolist():
                        if info.is_dir():
                            continue
                        child = info.filename
                        try:
                            child_data = zf.read(info)
                        except Exception:
                            stats["invalid"] += 1
                            continue
                        if child.lower().endswith(".zip"):
                            stats["nested_zips"] += 1
                        # Evitar nombres ambiguos entre ZIPs con una ruta legible.
                        display_name = f"{name} :: {child}" if depth > 0 or name.lower().endswith(".zip") else child
                        yield from walk(display_name, child_data, depth + 1)
            except zipfile.BadZipFile:
                stats["invalid"] += 1
            return

        stats["invalid"] += 1

    for item in uploaded_files:
        name = getattr(item, "name", "archivo")
        try:
            data = item.getvalue()
        except Exception:
            data = item.read()
        yield from walk(name, data, 0)

    return stats


def collect_jobs(uploaded_files) -> tuple[list[PDFJob], dict]:
    jobs: list[PDFJob] = []
    seen: dict[str, str] = {}
    duplicate_names = []
    stats = {
        "pdf_candidates": 0,
        "zip_files": sum(1 for f in uploaded_files if getattr(f, "name", "").lower().endswith(".zip")),
        "invalid": 0,
        "nested_zips": 0,
    }

    source_iter = iter_uploaded_sources(uploaded_files)
    for name, data in source_iter:
        stats["pdf_candidates"] += 1
        digest = sha256_bytes(data)
        if digest in seen:
            duplicate_names.append((name, seen[digest]))
            continue
        seen[digest] = name
        jobs.append(PDFJob(name=name, data=data, sha256=digest))

    stats["duplicates"] = len(duplicate_names)
    stats["duplicate_names"] = duplicate_names
    stats["unique_pdfs"] = len(jobs)
    return jobs, stats


# =========================
# EXTRACCIÓN DE UN PDF
# =========================
def empty_result_df() -> pd.DataFrame:
    return pd.DataFrame(columns=OUTPUT_COLUMNS)


def process_pdf(job: PDFJob, piece_number: int) -> dict:
    """Procesa un PDF completo. Diseñado para ejecutarse en un hilo."""
    rows_apoyos = []
    rows_levas = []
    rows_chatter_levas = []
    rows_chatter_apoyos = []
    warning = None

    try:
        with pdfplumber.open(io.BytesIO(job.data)) as pdf:
            if not pdf.pages:
                return {"name": job.name, "piece": piece_number, "error": "PDF sin páginas"}
            first_page_text = pdf.pages[0].extract_text() or ""
            second_page_text = pdf.pages[1].extract_text() if len(pdf.pages) > 1 else ""
    except Exception as exc:
        return {"name": job.name, "piece": piece_number, "error": f"No se pudo leer PDF: {exc}"}

    main_journals_start = first_page_text.find("MAIN JOURNALS:")
    lobes_start = first_page_text.find("LOBES:")
    chatter_start = first_page_text.find("CHATTER:")

    # ---------- MAIN JOURNALS / APOYOS ----------
    if main_journals_start != -1 and lobes_start != -1:
        main_text = first_page_text[main_journals_start:lobes_start]
        headers_pdf = [
            "Measured Diameter", "Error", "Roundness", "Runout", "Concentricity",
            "Parallelism", "Taper", "Cylindricity"
        ]
        char_mapping = {
            "Error": "Diametro", "Roundness": "Roundness", "Runout": "Runout",
            "Concentricity": "Concentricity", "Parallelism": "Parallelism",
            "Taper": "Taper", "Cylindricity": "Cylindricity"
        }
        pattern = re.compile(
            r"(" + "|".join(re.escape(i) for i in APOYO_IDENTIFIERS) +
            r")\s+([\s\d.\-#]+(?:\s+J[\d\-]+:\s+[\d.\-]+)?(?:\s+L[\d\-]+:\s+[\d.\-]+)?(?:\s+Tol:\s+[\d.\-]+)?)"
        )
        for line in main_text.split("\n"):
            match = pattern.match(line.strip())
            if not match:
                continue
            apoyo_val = match.group(1).strip()
            values_str = match.group(2).strip()
            raw_values = [
                v for v in re.split(r"\s+", values_str)
                if v and not re.match(r"J\d-\d:|J\d:", v) and v != "Tol:"
            ]
            extracted = {char: None for char in CHARACTERISTICS_FOR_COLUMNS}

            if apoyo_val == "1:A L":
                if len(raw_values) >= 6:
                    extracted["Diametro"] = raw_values[1]
                    extracted["Roundness"] = raw_values[2]
                    extracted["Runout"] = raw_values[3]
                    extracted["Concentricity"] = raw_values[4]
                    extracted["Taper"] = raw_values[5]
            elif apoyo_val == "5:B":
                if len(raw_values) >= 7:
                    extracted["Diametro"] = raw_values[1]
                    extracted["Roundness"] = raw_values[2]
                    extracted["Runout"] = raw_values[3]
                    extracted["Concentricity"] = raw_values[4]
                    extracted["Taper"] = raw_values[5]
                    extracted["Cylindricity"] = raw_values[6]
            else:
                idx = 0
                for header in headers_pdf:
                    if header == "Measured Diameter":
                        idx += 1
                        continue
                    if idx >= len(raw_values):
                        break
                    if header in char_mapping:
                        extracted[char_mapping[header]] = raw_values[idx]
                    idx += 1

            for char_name, measurement in extracted.items():
                if measurement is None:
                    continue
                rows_apoyos.append({
                    "Nombre del archivo": job.name,
                    "Pieza": piece_number,
                    "Leva": None,
                    "Apoyo": apoyo_val,
                    "Caracteristica": char_name,
                    "Medicion": str(measurement).replace("#", ""),
                    "Area": "Apoyos",
                    "Resultado": get_resultado(measurement),
                })

    # ---------- LOBES / LEVAS ----------
    if lobes_start != -1:
        lobes_text = first_page_text[lobes_start:chatter_start if chatter_start != -1 else len(first_page_text)]
        line_pattern = re.compile(r"^(?:\d+:)?\s*([A-Z0-9-:]+)\s+(.*)")
        lines = lobes_text.split("\n")
        start_idx = next((i for i, line in enumerate(lines) if line.strip().startswith("1: EXH-1")), -1)
        if start_idx != -1:
            for line in lines[start_idx:]:
                match = line_pattern.match(line.strip())
                if not match:
                    continue
                leva_val = match.group(1).strip()
                raw_values = [v for v in re.split(r"\s+", match.group(2).strip()) if v]
                if len(raw_values) == len(LOBES_HEADERS):
                    processed = raw_values
                elif len(raw_values) == 2 * len(LOBES_HEADERS):
                    processed = raw_values[::2]
                else:
                    continue
                for char_name, measurement in zip(LOBES_HEADERS, processed):
                    rows_levas.append({
                        "Nombre del archivo": job.name,
                        "Pieza": piece_number,
                        "Leva": leva_val,
                        "Apoyo": None,
                        "Caracteristica": char_name,
                        "Medicion": str(measurement).replace("#", ""),
                        "Area": "Levas",
                        "Resultado": get_resultado(measurement),
                    })

    # ---------- CHATTER LEVAS ----------
    if chatter_start != -1:
        chatter_text = first_page_text[chatter_start:]
        chatter_lines = chatter_text.split("\n")
        data_start = next((i for i, line in enumerate(chatter_lines) if re.match(r"^\s*\d+:", line.strip())), -1)
        if data_start != -1:
            for line in chatter_lines[data_start:]:
                match = re.match(r"^(\d+):\s+(.*)", line.strip())
                if not match:
                    continue
                leva_val = match.group(1).strip()
                raw_values = [v for v in re.split(r"\s+", match.group(2).strip()) if v]
                if len(raw_values) != len(CHATTER_LEVAS_CHARACTERISTICS) * 4:
                    continue
                n = len(CHATTER_LEVAS_CHARACTERISTICS)
                amplitudes_bc = [raw_values[i * 2] for i in range(n)]
                amplitudes_la = [raw_values[n * 2 + i * 2] for i in range(n)]
                for char_name, measurement in zip(CHATTER_LEVAS_CHARACTERISTICS, amplitudes_bc):
                    rows_chatter_levas.append({
                        "Nombre del archivo": job.name,
                        "Pieza": piece_number,
                        "Leva": leva_val,
                        "Apoyo": None,
                        "Caracteristica": char_name,
                        "Medicion": str(measurement).replace("#", ""),
                        "Area": "Base Circle",
                        "Resultado": get_resultado(measurement, char_name, "Base Circle"),
                    })
                for char_name, measurement in zip(CHATTER_LEVAS_CHARACTERISTICS, amplitudes_la):
                    rows_chatter_levas.append({
                        "Nombre del archivo": job.name,
                        "Pieza": piece_number,
                        "Leva": leva_val,
                        "Apoyo": None,
                        "Caracteristica": char_name,
                        "Medicion": str(measurement).replace("#", ""),
                        "Area": "Lift Area",
                        "Resultado": get_resultado(measurement, char_name, "Lift Area"),
                    })

    # ---------- CHATTER APOYOS ----------
    if second_page_text:
        journal_match = re.search(r"CHATTER:\s*-+\s*Journals\s*-+", second_page_text)
        if journal_match:
            journal_text = second_page_text[journal_match.start():]
            lines = journal_text.split("\n")
            data_start = next((i for i, line in enumerate(lines) if re.match(r"^\s*\d+:", line.strip())), -1)
            if data_start != -1:
                for line in lines[data_start:]:
                    match = re.match(r"^(\d+):\s+(.*)", line.strip())
                    if not match:
                        continue
                    original = match.group(1).strip()
                    apoyo = APOYO_RENAME_MAPPING.get(original, original)
                    raw_values = [v for v in re.split(r"\s+", match.group(2).strip()) if v and v != "@UPR"]
                    n_amp = len(raw_values) // 2
                    for i in range(min(n_amp, len(CHATTER_APOYOS_CHARACTERISTICS))):
                        char_name = CHATTER_APOYOS_CHARACTERISTICS[i]
                        measurement = raw_values[i * 2]
                        rows_chatter_apoyos.append({
                            "Nombre del archivo": job.name,
                            "Pieza": piece_number,
                            "Leva": None,
                            "Apoyo": apoyo,
                            "Caracteristica": char_name,
                            "Medicion": str(measurement).replace("#", ""),
                            "Area": "Apoyos",
                            "Resultado": get_resultado(measurement, char_name),
                        })

    def make_df(rows):
        df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
        if not df.empty:
            df["Medicion"] = pd.to_numeric(df["Medicion"], errors="coerce")
        return df

    return {
        "name": job.name,
        "piece": piece_number,
        "error": None,
        "apoyos": make_df(rows_apoyos),
        "levas": make_df(rows_levas),
        "chatter_levas": make_df(rows_chatter_levas),
        "chatter_apoyos": make_df(rows_chatter_apoyos),
        "sha256": job.sha256,
        "warning": warning,
    }


# =========================
# CONSOLIDACIÓN / ANÁLISIS
# =========================
def concat_results(results: list[dict], key: str) -> pd.DataFrame:
    frames = [r[key] for r in results if isinstance(r.get(key), pd.DataFrame) and not r[key].empty]
    return pd.concat(frames, ignore_index=True) if frames else empty_result_df()


def build_analysis(df_master: pd.DataFrame, df_nok: pd.DataFrame):
    if df_master.empty:
        return pd.DataFrame(), pd.DataFrame()

    total_pieces = int(df_master["Pieza"].nunique())
    nok_pieces = df_nok["Pieza"].unique() if not df_nok.empty else np.array([])
    num_nok_pieces = len(nok_pieces)
    num_ok_pieces = total_pieces - num_nok_pieces
    total_chars = len(df_master)
    ok_chars = int((df_master["Resultado"] == "Ok").sum())
    nok_chars = int((df_master["Resultado"] == "Nok").sum())
    rejection_pct = (num_nok_pieces / total_pieces * 100) if total_pieces else 0
    ftq = (num_ok_pieces / total_pieces * 100) if total_pieces else 0

    specific = "(5) 301-400 UPR"
    specific_nok_pieces = df_nok.loc[df_nok["Caracteristica"] == specific, "Pieza"].unique() if not df_nok.empty else []
    only_specific = []
    for piece in specific_nok_pieces:
        chars = df_nok.loc[df_nok["Pieza"] == piece, "Caracteristica"].unique()
        if len(chars) == 1 and chars[0] == specific:
            only_specific.append(piece)
    only_df = df_nok[(df_nok["Pieza"].isin(only_specific)) & (df_nok["Caracteristica"] == specific)]

    analysis = pd.DataFrame({
        "Metrica": [
            "Total de piezas analizadas",
            "Piezas OK (todas las características OK)",
            "Piezas NOK (al menos una característica NOK)",
            "Porcentaje de Rechazo",
            "FTQ (First Time Quality)",
            "Total de características analizadas",
            "Características OK",
            "Características NOK",
            "% Características OK",
            "% Características NOK",
            f'Piezas rechazadas SOLO por "{specific}"',
            f'Piezas rechazadas por "{specific}" y alguna otra',
            f'Piezas SOLO por "{specific}" en Lift Area',
            f'Piezas SOLO por "{specific}" en Base Circle',
        ],
        "Valor": [
            total_pieces, num_ok_pieces, num_nok_pieces,
            f"{rejection_pct:.2f}%", f"{ftq:.2f}%",
            total_chars, ok_chars, nok_chars,
            f"{ok_chars / total_chars * 100:.2f}%" if total_chars else "0.00%",
            f"{nok_chars / total_chars * 100:.2f}%" if total_chars else "0.00%",
            len(only_specific),
            len([p for p in specific_nok_pieces if p not in only_specific]),
            only_df.loc[only_df["Area"] == "Lift Area", "Pieza"].nunique(),
            only_df.loc[only_df["Area"] == "Base Circle", "Pieza"].nunique(),
        ],
    })

    counts = df_nok["Caracteristica"].value_counts().reset_index() if not df_nok.empty else pd.DataFrame(columns=["Caracteristica", "Cantidad_NOK"])
    if not counts.empty:
        counts.columns = ["Caracteristica", "Cantidad_NOK"]
        top = counts.head(15)["Caracteristica"].tolist()
    else:
        top = []

    detailed = []
    for char in top:
        all_c = df_master[df_master["Caracteristica"] == char]
        nok_c = df_nok[df_nok["Caracteristica"] == char]
        detailed.append({
            "Caracteristica": char,
            "Cantidad Piezas NOK": int(nok_c["Pieza"].nunique()),
            "Promedio Total Med.": all_c["Medicion"].mean(),
            "Std Total Med.": all_c["Medicion"].std(),
            "Promedio NOK Med.": nok_c["Medicion"].mean(),
            "Std NOK Med.": nok_c["Medicion"].std(),
            "Max NOK Med.": nok_c["Medicion"].max(),
            "Min NOK Med.": nok_c["Medicion"].min(),
        })
    detailed_df = pd.DataFrame(detailed)
    return analysis, detailed_df


def prepare_workbooks(dfs: dict[str, pd.DataFrame]) -> list[tuple[str, bytes]]:
    """Genera uno o más XLSX respetando el límite de filas de Excel."""
    parts: dict[str, list[pd.DataFrame]] = {}
    for sheet, df in dfs.items():
        if df is None or df.empty:
            continue
        if len(df) <= MAX_EXCEL_ROWS:
            parts[sheet] = [df]
        else:
            parts[sheet] = [df.iloc[i:i + MAX_EXCEL_ROWS] for i in range(0, len(df), MAX_EXCEL_ROWS)]

    if not parts:
        return []

    # Si alguna hoja se divide, todos los fragmentos se mantienen en el mismo workbook
    # mientras cada hoja individual respete el límite. Si el workbook crece demasiado,
    # Streamlit sigue entregando un único archivo para facilitar al usuario.
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for sheet, chunks in parts.items():
            for idx, chunk in enumerate(chunks, start=1):
                sheet_name = sheet if len(chunks) == 1 else f"{sheet[:28]}_{idx}"
                chunk.to_excel(writer, index=False, sheet_name=sheet_name[:31])
                ws = writer.book[sheet_name[:31]]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
    return [("Reporte_NoK.xlsx", bio.getvalue())]


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(
    page_title="Reporte NoK | Procesamiento masivo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 Reporte NoK")
st.caption(f"Procesamiento masivo de PDFs y ZIPs · {APP_VERSION}")

with st.sidebar:
    st.header("⚙️ Configuración")
    workers = st.slider(
        "Trabajadores paralelos",
        min_value=2,
        max_value=32,
        value=DEFAULT_WORKERS,
        help="Más trabajadores no siempre significa mayor velocidad. Para PDFs normalmente 6–16 es un buen rango."
    )
    st.info(
        f"Umbral Chatter Levas **(5) 301-400 UPR: >= {CHATTER_LEVAS_THRESHOLD:.4f} → NOK**\n\n"
        f"Chatter Apoyos: **{CHATTER_APOYOS_THRESHOLD:.5f}** cuando aplique.\n\n"
        "El símbolo # siempre es NOK."
    )
    st.markdown("**Entrada aceptada:** PDF, ZIP y combinación PDF + ZIP.")
    st.markdown(f"**ZIP anidados:** hasta {MAX_ZIP_DEPTH} niveles.")

st.subheader("1. Carga de archivos")
with st.form("upload_form", clear_on_submit=False):
    uploaded_files = st.file_uploader(
        "Selecciona PDFs individuales, ZIPs o ambos",
        type=["pdf", "zip"],
        accept_multiple_files=True,
        help="Puedes seleccionar miles de PDFs y uno o varios ZIP. Para grandes volúmenes, selecciona todos y después pulsa Procesar.",
    )
    submitted = st.form_submit_button("🚀 Analizar archivos", type="primary", use_container_width=True)

if uploaded_files:
    total_input_mb = sum(getattr(f, "size", 0) for f in uploaded_files) / (1024 ** 2)
    c1, c2, c3 = st.columns(3)
    c1.metric("Archivos seleccionados", f"{len(uploaded_files):,}")
    c2.metric("Tamaño de carga", f"{total_input_mb:,.1f} MB")
    c3.metric("Modo", "PDF + ZIP" if any(f.name.lower().endswith('.zip') for f in uploaded_files) and any(f.name.lower().endswith('.pdf') for f in uploaded_files) else ("ZIP" if any(f.name.lower().endswith('.zip') for f in uploaded_files) else "PDF"))

if submitted:
    if not uploaded_files:
        st.warning("Selecciona al menos un PDF o ZIP.")
        st.stop()

    total_start = time.perf_counter()
    status = st.status("Preparando archivos...", expanded=True)
    progress = st.progress(0, text="Leyendo PDFs y ZIPs...")

    try:
        jobs, input_stats = collect_jobs(uploaded_files)
    except Exception as exc:
        status.update(label="Error preparando archivos", state="error")
        st.exception(exc)
        st.stop()

    status.write(f"PDFs candidatos encontrados: {input_stats['pdf_candidates']:,}")
    status.write(f"PDFs únicos: {input_stats['unique_pdfs']:,}")
    status.write(f"Duplicados eliminados: {input_stats['duplicates']:,}")
    status.write(f"ZIPs detectados: {input_stats['zip_files']:,}")

    if not jobs:
        status.update(label="No se encontraron PDFs válidos", state="error")
        st.error("No se encontró ningún PDF procesable en la selección.")
        st.stop()

    # Orden determinista para que Pieza sea reproducible.
    jobs.sort(key=lambda j: j.name.lower())
    total = len(jobs)
    results = []
    errors = []
    completed = 0
    process_start = time.perf_counter()

    status.write(f"Iniciando {workers} trabajadores para {total:,} PDFs...")
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(process_pdf, job, idx): idx
            for idx, job in enumerate(jobs, start=1)
        }
        for future in as_completed(futures):
            completed += 1
            try:
                result = future.result()
                results.append(result)
                if result.get("error"):
                    errors.append({"Pieza": result["piece"], "Nombre del archivo": result["name"], "Error": result["error"]})
            except Exception as exc:
                idx = futures[future]
                errors.append({"Pieza": idx, "Nombre del archivo": jobs[idx - 1].name, "Error": str(exc)})
            elapsed = time.perf_counter() - process_start
            rate = completed / elapsed if elapsed > 0 else 0
            eta = (total - completed) / rate if rate > 0 else 0
            progress.progress(completed / total, text=f"Procesados {completed:,}/{total:,} · {rate:.1f} PDF/s · ETA {eta/60:.1f} min")

    results.sort(key=lambda r: r.get("piece", 0))
    status.write(f"Procesamiento terminado en {(time.perf_counter() - process_start):.1f} s.")

    df_apoyos = concat_results(results, "apoyos")
    df_levas = concat_results(results, "levas")
    df_chatter_levas = concat_results(results, "chatter_levas")
    df_chatter_apoyos = concat_results(results, "chatter_apoyos")

    # Chatter Journal Ford derivado
    df_chatter_ford = df_chatter_apoyos[df_chatter_apoyos["Caracteristica"].isin(CHATTER_FORD_ORIGINAL)].copy() if not df_chatter_apoyos.empty else empty_result_df()
    if not df_chatter_ford.empty:
        df_chatter_ford["Caracteristica"] = df_chatter_ford["Caracteristica"].replace(CHATTER_FORD_RENAME)
        df_chatter_ford["Resultado"] = [
            get_resultado_chatter_journal_ford(v, c)
            for v, c in zip(df_chatter_ford["Medicion"], df_chatter_ford["Caracteristica"])
        ]

    df_master = pd.concat(
        [df_apoyos, df_levas, df_chatter_levas, df_chatter_apoyos],
        ignore_index=True
    ) if any(not x.empty for x in [df_apoyos, df_levas, df_chatter_levas, df_chatter_apoyos]) else empty_result_df()
    df_nok = df_master[df_master["Resultado"] == "Nok"].copy() if not df_master.empty else empty_result_df()
    df_analysis, df_detailed = build_analysis(df_master, df_nok)

    errors_df = pd.DataFrame(errors)

    # Guardar resultados en sesión para no perderlos en reruns de visualización.
    st.session_state["report_data"] = {
        "apoyos": df_apoyos,
        "levas": df_levas,
        "chatter_levas": df_chatter_levas,
        "chatter_apoyos": df_chatter_apoyos,
        "chatter_ford": df_chatter_ford,
        "master": df_master,
        "nok": df_nok,
        "analysis": df_analysis,
        "detailed": df_detailed,
        "errors": errors_df,
        "input_stats": input_stats,
        "elapsed": time.perf_counter() - total_start,
    }
    status.update(label="✅ Reporte terminado", state="complete", expanded=False)
    st.rerun()

# =========================
# VISUALIZACIÓN DEL REPORTE
# =========================
data = st.session_state.get("report_data")
if data:
    st.divider()
    st.subheader("2. Resultado del análisis")
    df_master = data["master"]
    df_nok = data["nok"]
    df_analysis = data["analysis"]

    if df_analysis.empty:
        st.warning("No se extrajeron datos de los PDFs.")
    else:
        metrics = {row["Metrica"]: row["Valor"] for _, row in df_analysis.iterrows()}
        m = st.columns(5)
        m[0].metric("Piezas", f"{int(metrics.get('Total de piezas analizadas', 0)):,}")
        m[1].metric("Piezas NOK", f"{int(metrics.get('Piezas NOK (al menos una característica NOK)', 0)):,}")
        m[2].metric("Piezas OK", f"{int(metrics.get('Piezas OK (todas las características OK)', 0)):,}")
        m[3].metric("Rechazo", metrics.get("Porcentaje de Rechazo", "0.00%"))
        m[4].metric("FTQ", metrics.get("FTQ (First Time Quality)", "0.00%"))

        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Resumen", "🚨 NOK", "📊 Características", "📋 Datos", "⚠️ Errores"])

        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("#### OK vs NOK por pieza")
                chart_pieces = pd.DataFrame({
                    "Resultado": ["OK", "NOK"],
                    "Piezas": [
                        int(metrics.get("Piezas OK (todas las características OK)", 0)),
                        int(metrics.get("Piezas NOK (al menos una característica NOK)", 0)),
                    ],
                }).set_index("Resultado")
                st.bar_chart(chart_pieces)
            with c2:
                st.markdown("#### NOK por área")
                if not df_nok.empty:
                    area_chart = df_nok.groupby("Area")["Pieza"].nunique().sort_values(ascending=False).to_frame("Piezas NOK")
                    st.bar_chart(area_chart)
                else:
                    st.info("No hay NOK.")
            st.markdown("#### Resumen completo")
            st.dataframe(df_analysis, use_container_width=True, hide_index=True)

        with tab2:
            st.markdown("#### Características con más NOK")
            if df_nok.empty:
                st.success("No se encontraron resultados NOK.")
            else:
                nok_by_char = df_nok.groupby("Caracteristica")["Pieza"].nunique().sort_values(ascending=False).head(15).to_frame("Piezas NOK")
                st.bar_chart(nok_by_char)
                st.dataframe(
                    df_nok.sort_values(["Pieza", "Area", "Caracteristica"]),
                    use_container_width=True,
                    height=500,
                    hide_index=True,
                )

        with tab3:
            st.markdown("#### Análisis estadístico de las principales características NOK")
            st.dataframe(data["detailed"], use_container_width=True, hide_index=True)
            if not df_nok.empty:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**NOK por área**")
                    st.bar_chart(df_nok.groupby("Area").size().to_frame("NOK"))
                with c2:
                    st.markdown("**NOK por característica**")
                    st.bar_chart(df_nok.groupby("Caracteristica").size().sort_values(ascending=False).head(15).to_frame("NOK"))

        with tab4:
            st.caption("Los datos completos se entregan en Excel. Aquí se muestra una vista limitada para mantener la interfaz rápida.")
            dataset_choice = st.selectbox("Hoja a visualizar", ["Apoyos", "Levas", "chatter Levas", "Chatter apoyos", "Chatter Journal Ford"])
            mapping = {
                "Apoyos": data["apoyos"], "Levas": data["levas"],
                "chatter Levas": data["chatter_levas"],
                "Chatter apoyos": data["chatter_apoyos"],
                "Chatter Journal Ford": data["chatter_ford"],
            }
            chosen = mapping[dataset_choice]
            st.write(f"Filas: {len(chosen):,}")
            st.dataframe(chosen.head(5000), use_container_width=True, height=550, hide_index=True)

        with tab5:
            if data["errors"].empty:
                st.success("Todos los PDFs terminaron sin errores de lectura/procesamiento.")
            else:
                st.warning(f"{len(data['errors']):,} PDFs presentaron errores.")
                st.dataframe(data["errors"], use_container_width=True, hide_index=True)

        # ---------- EXPORTACIÓN ----------
        st.divider()
        st.subheader("3. Descargar reporte")
        workbook_dfs = {
            "Apoyos": data["apoyos"],
            "Levas": data["levas"],
            "chatter Levas": data["chatter_levas"],
            "Chatter apoyos": data["chatter_apoyos"],
            "Chatter Journal Ford": data["chatter_ford"],
            "Nok": data["nok"],
            "Analisis_Resumen": data["analysis"],
            "Analisis_Detallado_NOK": data["detailed"],
            "Errores": data["errors"],
        }
        with st.spinner("Generando Excel..."):
            files_out = prepare_workbooks(workbook_dfs)
        for filename, content in files_out:
            st.download_button(
                label=f"⬇️ Descargar {filename}",
                data=content,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.caption(
            f"Tiempo total de ejecución: {data['elapsed']:.1f} s · "
            f"PDFs únicos: {data['input_stats']['unique_pdfs']:,} · "
            f"Duplicados eliminados: {data['input_stats']['duplicates']:,}"
        )
else:
    st.info("Selecciona PDFs/ZIPs arriba y pulsa **🚀 Analizar archivos**.")

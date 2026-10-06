# -*- coding: utf-8 -*-
"""
Reporte NoK Streamlit V9.3 - optimizado para Streamlit Community Cloud

Objetivos:
- PDFs individuales y uno/muchos ZIP.
- Mezcla PDF + ZIP.
- ZIP anidados hasta 8 niveles.
- Detección de duplicados por SHA-256.
- Procesamiento por lotes para controlar RAM.
- PDFs temporales en disco; no se conservan todos los bytes en memoria.
- Procesamiento paralelo limitado.
- Dashboard interactivo para Levas, Apoyos y características.
- Excel final generado con openpyxl en modo write-only para reducir RAM.

NOTA: Streamlit Community Cloud sigue manteniendo los archivos seleccionados por
el uploader en memoria. Por eso esta versión recomienda lotes de 100-150 PDFs.
Para ZIP grandes, se procesa su contenido directamente desde disco temporal.
"""

from __future__ import annotations

import hashlib
import io
import json
import os
import re
import shutil
import sqlite3
import tempfile
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Tuple

import numpy as np
import pandas as pd
import pdfplumber
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# PDF ejecutivo del dashboard
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak


# ============================================================
# CONFIGURACIÓN
# ============================================================
APP_VERSION = "V9.5"
MAX_ZIP_DEPTH = 8
MAX_INPUT_PDFS = 1500
PROCESS_CHUNK_SIZE = 150
MAX_WORKERS = 8

CHATTER_LEVAS_THRESHOLD = 0.0001
CHATTER_APOYOS_THRESHOLD = 0.00008

CHARACTERISTICS_FOR_APOYOS_COLUMNS = [
    "Diametro", "Roundness", "Runout", "Concentricity",
    "Parallelism", "Taper", "Cylindricity"
]

APOYO_IDENTIFIERS = ["Aux:G", "1:A L", "1:A R", "2:C", "3:D", "4:E", "5:B"]

APOYO_UNIFICADO_MAPPING = {
    "Aux:G": "Aux:G",
    "1:A L": "1",
    "1:A R": "2",
    "2:C": "3",
    "3:D": "4",
    "4:E": "5",
    "5:B": "6",
}

CHATTER_APOYOS_CHARACTERISTICS = [
    "(1) 5 - 8 UPR", "(2) 9 - 15 UPR", "(3) 16 - 23 UPR",
    "(4) 24 - 28 UPR", "(5) 29- 45 UPR", "(6) 46-70 UPR",
    "(7) 71-140 UPR", "(8) 141-215 UPR", "(9) 216-270 UPR",
    "(10) 271-300 UPR",
]

APOYO_RENAME_MAPPING = {f"{i}:": str(i) for i in range(1, 7)}

CHATTER_CHARACTERISTICS = [
    "(1) 40- 80 UPR", "(2) 81-140 UPR", "(3) 141-190 UPR",
    "(4) 191-300 UPR", "(5) 301-400 UPR",
]

LOBES_HEADERS = [
    "AngleErr", "BC-Rad'sErr", "BC-Runout", "BC-Vel./10°",
    "Ramp-MaxLift", "Nose-MaxLift", "Ramp+9°Vel/1°",
    "Nose-Vel./1°", "Taper", "Center-Dev",
]

COLUMNS = [
    "Nombre del archivo", "Pieza", "Leva", "Apoyo", "Caracteristica",
    "Medicion", "Area", "Resultado", "Spec Min", "Spec Max",
]

SHEET_ORDER = [
    "Apoyos", "Levas", "chatter Levas", "Chatter apoyos",
    "Chatter Journal Ford", "Master", "Nok", "Analisis_Resumen",
    "Analisis_Detallado_NOK", "Errores",
]

SPECS = {
    "Aux:G": {
        "Diametro": (0.0, -0.0130, 0.0130),
        "Roundness": (0.0, 0.0, 0.0120),
        "Runout": (0.0, 0.0, 0.0500),
        "Concentricity": (0.0, 0.0, 0.0300),
        "Parallelism": (0.0, 0.0, 0.0150),
        "Taper": (0.0, -0.0150, 0.0150),
        "Cylindricity": (0.0, 0.0, 0.0300),
    },
    "Genericos_Apoyos": {
        "Diametro": (0.0, -0.0130, 0.0130),
        "Roundness": (0.0, 0.0, 0.0120),
        "Runout": (0.0, 0.0, 0.0300),
        "Concentricity": (0.0, 0.0, 0.0300),
        "Parallelism": (0.0, 0.0, 0.0150),
        "Taper": (0.0, -0.0150, 0.0150),
        "Cylindricity": (0.0, 0.0, 0.0300),
    },
    "Levas_Default": {
        "AngleErr": (0.0, -0.250, 0.250),
        "BC-Rad'sErr": (0.0, -0.0380, 0.0380),
        "BC-Runout": (0.0, 0.0, 0.0250),
        "BC-Vel./10°": (0.0, -0.0120, 0.0120),
        "Ramp-MaxLift": (0.0, -0.0310, 0.0310),
        "Nose-MaxLift": (0.0, -0.0250, 0.0250),
        "Ramp+9°Vel/1°": (0.0, -0.0130, 0.0130),
        "Nose-Vel./1°": (0.0, -0.0033, 0.0033),
        "Taper": (0.0, -0.0160, 0.0160),
        "Center-Dev": (0.0, -0.0080, 0.0080),
    },
    "Chatter_Levas": {
        "(1) 40- 80 UPR": (0.0, 0.0, 0.0003250),
        "(2) 81-140 UPR": (0.0, 0.0, 0.0001600),
        "(3) 141-190 UPR": (0.0, 0.0, 0.0001250),
        "(4) 191-300 UPR": (0.0, 0.0, 0.0000800),
        # Regla solicitada: en Levas, (5) 301-400 UPR = 0.0001
        "(5) 301-400 UPR": (0.0, 0.0, CHATTER_LEVAS_THRESHOLD),
    },
    "Chatter_Apoyos": {
        "(1) 5 - 8 UPR": (0.0, 0.0, 0.0010000),
        "(2) 9 - 15 UPR": (0.0, 0.0, 0.0007500),
        "(3) 16 - 23 UPR": (0.0, 0.0, 0.0009500),
        "(4) 24 - 28 UPR": (0.0, 0.0, 0.0009500),
        "(5) 29- 45 UPR": (0.0, 0.0, 0.0005850),
        "(6) 46-70 UPR": (0.0, 0.0, 0.0005800),
        "(7) 71-140 UPR": (0.0, 0.0, 0.0003000),
        "(8) 141-215 UPR": (0.0, 0.0, 0.0003000),
        "(9) 216-270 UPR": (0.0, 0.0, 0.0003000),
        "(10) 271-300 UPR": (0.0, 0.0, 0.0003000),
    },
}

FORD_THRESHOLDS = {
    "(1) 29 - 45 UPR": 0.00033,
    "(2) 46 -70 UPR": 0.00033,
    "(3) 71 - 140 UPR": 0.00016,
    "(4) 141 - 215 UPR": 0.00013,
    "(5) 216 - 270 UPR": 0.00008,
}

FORD_RENAME = {
    "(5) 29- 45 UPR": "(1) 29 - 45 UPR",
    "(6) 46-70 UPR": "(2) 46 -70 UPR",
    "(7) 71-140 UPR": "(3) 71 - 140 UPR",
    "(8) 141-215 UPR": "(4) 141 - 215 UPR",
    "(9) 216-270 UPR": "(5) 216 - 270 UPR",
}


# ============================================================
# STREAMLIT CONFIG / ESTILO
# ============================================================
st.set_page_config(
    page_title="Reporte NoK | V9",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main {padding-top: 1rem;}
    .hero {padding: 1.1rem 1.4rem; border-radius: 18px;
           background: linear-gradient(135deg,#111827,#1f2937 55%,#374151);
           color:white; margin-bottom:1rem;}
    .hero h1 {margin:0; font-size:2rem;}
    .hero p {margin:.35rem 0 0; opacity:.86;}
    .metric-card {padding: .8rem 1rem; border:1px solid #e5e7eb;
                  border-radius:14px; background:#fff;}
    .small-muted {color:#6b7280;font-size:.85rem;}
    .ok-pill {padding:.2rem .5rem;border-radius:999px;background:#ecfdf5;color:#047857;font-weight:700;}
    .nok-pill {padding:.2rem .5rem;border-radius:999px;background:#fef2f2;color:#b91c1c;font-weight:700;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div class='hero'>
      <h1>📊 Reporte NoK — Dashboard de Calidad</h1>
      <p>Procesamiento masivo de reportes | V{APP_VERSION.replace('V','')} | Levas + Apoyos + Chatter</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PERSISTENCIA TEMPORAL / SQLITE
# ============================================================
def init_workspace() -> Path:
    if "workspace" not in st.session_state:
        root = Path(tempfile.mkdtemp(prefix="reporte_nok_"))
        (root / "batches").mkdir(parents=True, exist_ok=True)
        st.session_state.workspace = str(root)
    root = Path(st.session_state.workspace)
    (root / "batches").mkdir(parents=True, exist_ok=True)
    return root


WORKSPACE = init_workspace()
DB_PATH = WORKSPACE / "manifest.sqlite"


def db_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """CREATE TABLE IF NOT EXISTS files (
            sha256 TEXT PRIMARY KEY,
            filename TEXT,
            piece INTEGER,
            status TEXT,
            error TEXT,
            processed_at TEXT
        )"""
    )
    conn.commit()
    return conn


def next_piece_number(conn) -> int:
    row = conn.execute("SELECT COALESCE(MAX(piece),0) FROM files").fetchone()
    return int(row[0]) + 1


def known_hash(conn, sha256: str) -> bool:
    return conn.execute("SELECT 1 FROM files WHERE sha256=?", (sha256,)).fetchone() is not None


def register_files_bulk(conn, rows):
    """Registra un bloque completo con un solo commit (más rápido que commit por PDF)."""
    if not rows:
        return
    conn.executemany(
        "INSERT OR REPLACE INTO files(sha256,filename,piece,status,error,processed_at) VALUES(?,?,?,?,?,?)",
        rows,
    )
    conn.commit()


def reset_workspace():
    old = st.session_state.get("workspace")
    if old:
        shutil.rmtree(old, ignore_errors=True)
    for k in ["workspace", "last_summary"]:
        st.session_state.pop(k, None)
    st.rerun()


# ============================================================
# PARSEO
# ============================================================
def numeric_value(value) -> Optional[float]:
    if value is None:
        return None
    s = str(value).replace(",", ".").replace("#", "").strip()
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def get_resultado(value_str, characteristic_name=None, area_type=None):
    """NOK if PDF has #, plus chatter threshold rules."""
    text_value = str(value_str)
    if "#" in text_value:
        return "Nok"

    val = numeric_value(text_value)
    if val is None:
        return "Ok"

    # Regla específica solicitada: Levas (5) 301-400 UPR, 0.0001.
    if characteristic_name == "(5) 301-400 UPR" and area_type == "Chatter_Levas":
        return "Nok" if val >= CHATTER_LEVAS_THRESHOLD else "Ok"

    return "Ok"


def get_resultado_chatter_journal_ford(medicion_val, characteristic_name):
    if "#" in str(medicion_val):
        return "Nok"
    value = numeric_value(medicion_val)
    if value is None:
        return "Nok"
    threshold = FORD_THRESHOLDS.get(characteristic_name)
    return "Nok" if threshold is not None and value > threshold else "Ok"


def spec_limits(characteristic: str, area_type: str, apoyo: Optional[str] = None):
    if area_type == "Apoyos":
        cfg = SPECS["Aux:G"] if apoyo == "Aux:G" else SPECS["Genericos_Apoyos"]
    elif area_type == "Levas":
        cfg = SPECS["Levas_Default"]
    elif area_type == "Chatter_Levas":
        cfg = SPECS["Chatter_Levas"]
    elif area_type == "Chatter_Apoyos":
        cfg = SPECS["Chatter_Apoyos"]
    else:
        return None, None
    if characteristic not in cfg:
        return None, None
    nom, neg, pos = cfg[characteristic]
    return nom + neg, nom + pos


def base_row(filename, piece, leva=None, apoyo=None, char=None, medicion=None, area=None, resultado=None, area_type=None):
    spec_min, spec_max = spec_limits(char, area_type or area, apoyo)
    return {
        "Nombre del archivo": filename,
        "Pieza": piece,
        "Leva": leva,
        "Apoyo": apoyo,
        "Caracteristica": char,
        "Medicion": numeric_value(medicion),
        "Area": area,
        "Resultado": resultado,
        "Spec Min": spec_min,
        "Spec Max": spec_max,
    }


def parse_pdf(pdf_path: Path, filename: str, piece: int) -> Dict[str, object]:
    result = {name: [] for name in ["Apoyos", "Levas", "chatter Levas", "Chatter apoyos"]}
    with pdfplumber.open(str(pdf_path)) as pdf:
        if not pdf.pages:
            raise ValueError("PDF sin páginas")
        first_page_text = pdf.pages[0].extract_text() or ""
        second_page_text = pdf.pages[1].extract_text() if len(pdf.pages) > 1 else ""

    main_journals_start = first_page_text.find("MAIN JOURNALS:")
    lobes_start = first_page_text.find("LOBES:")
    chatter_start = first_page_text.find("CHATTER:")

    # ---------------- MAIN JOURNALS / APOYOS ----------------
    if main_journals_start != -1 and lobes_start != -1:
        section = first_page_text[main_journals_start:lobes_start]
        pattern = re.compile(
            r"(" + "|".join(re.escape(i) for i in APOYO_IDENTIFIERS) +
            r")\s+([\s\d.\-#]+(?:\s+J[\d\-]+:\s+[\d.\-]+)?(?:\s+L[\d\-]+:\s+[\d.\-]+)?(?:\s+Tol:\s+[\d.\-]+)?)"
        )
        for line in section.split("\n"):
            match = pattern.match(line.strip())
            if not match:
                continue
            apoyo_original = match.group(1).strip()
            apoyo = APOYO_UNIFICADO_MAPPING.get(apoyo_original, apoyo_original)
            values_str = match.group(2).strip()
            raw = [v for v in re.split(r"\s+", values_str) if v and not re.match(r"J\d-\d:|J\d:", v) and v != "Tol:"]
            extracted = {char: None for char in CHARACTERISTICS_FOR_APOYOS_COLUMNS}

            if apoyo_original == "1:A L" and len(raw) >= 6:
                extracted.update({
                    "Diametro": raw[1], "Roundness": raw[2], "Runout": raw[3],
                    "Concentricity": raw[4], "Taper": raw[5]
                })
            elif apoyo_original == "5:B" and len(raw) >= 7:
                extracted.update({
                    "Diametro": raw[1], "Roundness": raw[2], "Runout": raw[3],
                    "Concentricity": raw[4], "Taper": raw[5], "Cylindricity": raw[6]
                })
            else:
                # El primer valor corresponde a Measured Diameter; los siguientes a las características.
                idx = 1
                for char in CHARACTERISTICS_FOR_APOYOS_COLUMNS:
                    if idx < len(raw):
                        extracted[char] = raw[idx]
                    idx += 1

            for char, measurement in extracted.items():
                if measurement is None:
                    continue
                result["Apoyos"].append(base_row(
                    filename, piece, apoyo=apoyo, char=char,
                    medicion=measurement, area="Apoyos",
                    resultado=get_resultado(measurement, char, "Apoyos"),
                    area_type="Apoyos"
                ))

    # ---------------- LOBES / LEVAS ----------------
    if lobes_start != -1:
        section = first_page_text[lobes_start:chatter_start if chatter_start != -1 else len(first_page_text)]
        pattern = re.compile(r"^\d+:\s*([A-Z0-9-:]+)\s+(.*)")
        for line in section.split("\n"):
            match = pattern.match(line.strip())
            if not match:
                continue
            leva = line.split(":")[0].strip()
            raw = [v for v in re.split(r"\s+", match.group(2).strip()) if v]
            if len(raw) == len(LOBES_HEADERS):
                processed = raw
            elif len(raw) == 2 * len(LOBES_HEADERS):
                processed = [raw[j] for j in range(0, len(raw), 2)]
            else:
                continue
            for i, char in enumerate(LOBES_HEADERS):
                measurement = processed[i]
                result["Levas"].append(base_row(
                    filename, piece, leva=leva, char=char,
                    medicion=measurement, area="Levas",
                    resultado=get_resultado(measurement, char, "Levas"),
                    area_type="Levas"
                ))

    # ---------------- CHATTER LEVAS ----------------
    if chatter_start != -1:
        section = first_page_text[chatter_start:]
        data_started = False
        for line in section.split("\n"):
            if not data_started and re.match(r"^\s*\d+:", line.strip()):
                data_started = True
            if not data_started:
                continue
            match = re.match(r"^(\d+):\s+(.*)", line.strip())
            if not match:
                continue
            leva = match.group(1).strip()
            raw = [v for v in re.split(r"\s+", match.group(2).strip()) if v]
            expected = len(CHATTER_CHARACTERISTICS) * 2 * 2
            if len(raw) != expected:
                continue
            n = len(CHATTER_CHARACTERISTICS)
            bc = [raw[i * 2] for i in range(n)]
            la = [raw[n * 2 + i * 2] for i in range(n)]
            for i, char in enumerate(CHATTER_CHARACTERISTICS):
                for area, measurement in [("Base Circle", bc[i]), ("Lift Area", la[i])]:
                    result["chatter Levas"].append(base_row(
                        filename, piece, leva=leva, char=char,
                        medicion=measurement, area=area,
                        resultado=get_resultado(measurement, char, "Chatter_Levas"),
                        area_type="Chatter_Levas"
                    ))

    # ---------------- CHATTER APOYOS ----------------
    if second_page_text:
        m = re.search(r"CHATTER:\s*-+\s*Journals\s*-+", second_page_text)
        if m:
            section = second_page_text[m.start():]
            data_started = False
            for line in section.split("\n"):
                if not data_started and re.match(r"^\s*\d+:", line.strip()):
                    data_started = True
                if not data_started:
                    continue
                match = re.match(r"^(\d+):\s+(.*)", line.strip())
                if not match:
                    continue
                apoyo_original = match.group(1).strip()
                apoyo = APOYO_RENAME_MAPPING.get(apoyo_original + ":", apoyo_original)
                raw = [v for v in re.split(r"\s+", match.group(2).strip()) if v and v != "@UPR"]
                num = len(raw) // 2
                for i in range(min(num, len(CHATTER_APOYOS_CHARACTERISTICS))):
                    char = CHATTER_APOYOS_CHARACTERISTICS[i]
                    measurement = raw[i * 2]
                    result["Chatter apoyos"].append(base_row(
                        filename, piece, apoyo=apoyo, char=char,
                        medicion=measurement, area="Apoyos",
                        resultado=get_resultado(measurement, char, "Chatter_Apoyos"),
                        area_type="Chatter_Apoyos"
                    ))

    # Ford view is derived later.
    for key in result:
        if result[key]:
            result[key] = pd.DataFrame(result[key], columns=COLUMNS)
        else:
            result[key] = pd.DataFrame(columns=COLUMNS)
    return result


# ============================================================
# ENTRADAS: PDF / ZIP / ZIP ANIDADO
# ============================================================
def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def safe_name(name: str) -> str:
    return Path(name).name.replace("/", "_").replace("\\", "_")


def extract_zip_recursive(zip_path: Path, dest_root: Path, depth=0) -> Iterator[Path]:
    if depth > MAX_ZIP_DEPTH:
        raise ValueError(f"Se excedió la profundidad máxima de ZIP ({MAX_ZIP_DEPTH}).")
    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = safe_name(info.filename)
            out = dest_root / f"d{depth}_{hashlib.sha1(info.filename.encode('utf-8','ignore')).hexdigest()[:8]}_{name}"
            with zf.open(info, "r") as src, out.open("wb") as dst:
                shutil.copyfileobj(src, dst, length=1024 * 1024)
            suffix = out.suffix.lower()
            if suffix == ".pdf":
                yield out
            elif suffix == ".zip":
                yield from extract_zip_recursive(out, dest_root, depth + 1)
                try:
                    out.unlink()
                except OSError:
                    pass


def materialize_uploaded_sources(uploaded_files: List, source_dir: Path) -> Tuple[List[Path], List[str]]:
    """Escribe cada upload al disco; devuelve PDFs listos y errores de entrada."""
    pdf_paths = []
    errors = []
    for idx, up in enumerate(uploaded_files):
        try:
            name = getattr(up, "name", f"archivo_{idx}")
            out = source_dir / f"upload_{idx}_{safe_name(name)}"
            with out.open("wb") as f:
                shutil.copyfileobj(up, f, length=1024 * 1024)
            if out.suffix.lower() == ".pdf":
                pdf_paths.append(out)
            elif out.suffix.lower() == ".zip":
                pdf_paths.extend(list(extract_zip_recursive(out, source_dir)))
            else:
                errors.append(f"Formato no soportado: {name}")
        except Exception as e:
            errors.append(f"{getattr(up,'name','archivo')}: {e}")
    return pdf_paths, errors


# ============================================================
# PROCESAMIENTO POR LOTES
# ============================================================
def process_pdf_job(args):
    path, filename, piece = args
    try:
        data = parse_pdf(path, filename, piece)
        return {"ok": True, "path": str(path), "filename": filename, "piece": piece, "data": data, "error": ""}
    except Exception as e:
        return {"ok": False, "path": str(path), "filename": filename, "piece": piece, "data": None, "error": f"{type(e).__name__}: {e}"}


def make_ford_df(chatter_apoyos: pd.DataFrame) -> pd.DataFrame:
    if chatter_apoyos.empty:
        return pd.DataFrame(columns=COLUMNS)
    df = chatter_apoyos[chatter_apoyos["Caracteristica"].isin(FORD_RENAME)].copy()
    if df.empty:
        return pd.DataFrame(columns=COLUMNS)
    df["Caracteristica"] = df["Caracteristica"].replace(FORD_RENAME)
    df["Resultado"] = [get_resultado_chatter_journal_ford(v, c) for v, c in zip(df["Medicion"], df["Caracteristica"])]
    df["Spec Min"] = 0.0
    df["Spec Max"] = df["Caracteristica"].map(FORD_THRESHOLDS)
    return df[COLUMNS]


def save_batch_frames(batch_dir: Path, frames: Dict[str, pd.DataFrame]):
    for sheet, df in frames.items():
        if df is None or df.empty:
            continue
        safe_sheet = re.sub(r"[^A-Za-z0-9_-]+", "_", sheet)
        path = batch_dir / f"{safe_sheet}.csv.gz"
        df.to_csv(path, index=False, compression={"method": "gzip", "compresslevel": 1})


def read_all_sheet(sheet: str) -> pd.DataFrame:
    files = sorted((WORKSPACE / "batches").glob(f"*/{re.sub(r'[^A-Za-z0-9_-]+','_',sheet)}.csv.gz"))
    frames = []
    for p in files:
        try:
            frames.append(pd.read_csv(p, compression="gzip"))
        except Exception:
            pass
    if not frames:
        return pd.DataFrame(columns=COLUMNS)
    return pd.concat(frames, ignore_index=True)


def all_processed_data() -> Dict[str, pd.DataFrame]:
    return {sheet: read_all_sheet(sheet) for sheet in ["Apoyos", "Levas", "chatter Levas", "Chatter apoyos", "Chatter Journal Ford"]}


def build_master(data: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    frames = [data.get(k, pd.DataFrame(columns=COLUMNS)) for k in ["Apoyos", "Levas", "chatter Levas", "Chatter apoyos"]]
    frames = [x for x in frames if not x.empty]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=COLUMNS)


def process_uploaded_batch(uploaded_files: List, workers: int) -> Dict[str, int]:
    conn = db_conn()
    source_dir = WORKSPACE / "incoming" / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    source_dir.mkdir(parents=True, exist_ok=True)

    with st.status("Preparando archivos…", expanded=True) as status:
        st.write(f"Recibidos: {len(uploaded_files)} archivo(s)")
        pdf_paths, input_errors = materialize_uploaded_sources(uploaded_files, source_dir)
        st.write(f"PDFs encontrados en PDF/ZIP: {len(pdf_paths)}")
        if input_errors:
            st.warning(f"{len(input_errors)} entrada(s) no se pudieron preparar.")

    unique_jobs = []
    duplicates = 0
    piece = next_piece_number(conn)
    seen_in_batch = set()

    for p in pdf_paths:
        try:
            digest = sha256_file(p)
            if digest in seen_in_batch or known_hash(conn, digest):
                duplicates += 1
                continue
            seen_in_batch.add(digest)
            unique_jobs.append((p, p.name, piece, digest))
            piece += 1
        except Exception as e:
            input_errors.append(f"{p.name}: SHA-256: {e}")

    if len(unique_jobs) > MAX_INPUT_PDFS:
        st.error(f"La carga contiene {len(unique_jobs):,} PDFs únicos. El máximo permitido por ejecución es {MAX_INPUT_PDFS:,} PDFs.")
        shutil.rmtree(source_dir, ignore_errors=True)
        return {"recibidos": len(uploaded_files), "pdfs": len(pdf_paths), "procesados": 0, "duplicados": duplicates, "errores": len(input_errors) + 1}

    batch_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    batch_dir = WORKSPACE / "batches" / batch_id
    batch_dir.mkdir(parents=True, exist_ok=True)

    errors = list(input_errors)
    progress = st.progress(0, text="Procesando PDFs…")
    started = time.perf_counter()
    completed = 0
    jobs = [(p, name, num, digest) for p, name, num, digest in unique_jobs]

    # Importante: el usuario puede cargar hasta 1,500 PDFs de una vez, pero el
    # motor sólo mantiene PROCESS_CHUNK_SIZE resultados PDF en memoria a la vez.
    for chunk_start in range(0, len(jobs), PROCESS_CHUNK_SIZE):
        chunk = jobs[chunk_start:chunk_start + PROCESS_CHUNK_SIZE]
        frames_acc = {k: [] for k in ["Apoyos", "Levas", "chatter Levas", "Chatter apoyos"]}
        chunk_errors = []
        registrations = []

        with st.status(f"Procesando bloque {chunk_start // PROCESS_CHUNK_SIZE + 1}…", expanded=False) as block_status:
            with ThreadPoolExecutor(max_workers=min(max(1, workers), max(1, len(chunk)))) as executor:
                future_map = {executor.submit(process_pdf_job, (p, name, num)): (p, name, num, digest) for p, name, num, digest in chunk}
                for future in as_completed(future_map):
                    p, name, num, digest = future_map[future]
                    try:
                        out = future.result()
                        if out["ok"]:
                            for sheet, df in out["data"].items():
                                if not df.empty:
                                    frames_acc[sheet].append(df)
                            registrations.append((digest, name, num, "OK", "", datetime.now().isoformat(timespec="seconds")))
                        else:
                            registrations.append((digest, name, num, "ERROR", out["error"], datetime.now().isoformat(timespec="seconds")))
                            chunk_errors.append(f"{name}: {out['error']}")
                    except Exception as e:
                        err = f"{type(e).__name__}: {e}"
                        registrations.append((digest, name, num, "ERROR", err, datetime.now().isoformat(timespec="seconds")))
                        chunk_errors.append(f"{name}: {err}")

                    completed += 1
                    elapsed = max(time.perf_counter() - started, 0.001)
                    rate = completed / elapsed
                    eta = (len(jobs) - completed) / rate if rate else 0
                    progress.progress(completed / max(len(jobs), 1), text=f"{completed:,}/{len(jobs):,} | {rate:.1f} PDF/s | ETA {eta/60:.1f} min")

            # Un solo commit por bloque evita miles de operaciones de disco SQLite.
            register_files_bulk(conn, registrations)

            # Escribimos inmediatamente el bloque en CSV comprimido y liberamos
            # sus DataFrames antes de comenzar el siguiente bloque.
            chunk_frames = {}
            for sheet, parts in frames_acc.items():
                if parts:
                    chunk_frames[sheet] = pd.concat(parts, ignore_index=True)
            ford = make_ford_df(chunk_frames.get("Chatter apoyos", pd.DataFrame(columns=COLUMNS)))
            if not ford.empty:
                chunk_frames["Chatter Journal Ford"] = ford
            save_batch_frames(batch_dir, chunk_frames)
            del chunk_frames, frames_acc
            errors.extend(chunk_errors)
            block_status.update(label=f"Bloque terminado: {len(chunk):,} PDFs", state="complete")

    (batch_dir / "errors.json").write_text(json.dumps(errors, ensure_ascii=False, indent=2), encoding="utf-8")
    (batch_dir / "meta.json").write_text(json.dumps({
        "batch_id": batch_id, "files": len(uploaded_files), "pdfs": len(pdf_paths),
        "processed": len(jobs), "duplicates": duplicates,
        "errors": len(errors), "elapsed_s": time.perf_counter() - started,
        "input_limit": MAX_INPUT_PDFS, "internal_chunk_size": PROCESS_CHUNK_SIZE,
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    # Los PDFs temporales se eliminan; sólo quedan resultados compactos CSV.gz.
    shutil.rmtree(source_dir, ignore_errors=True)

    st.success(f"Carga terminada: {len(jobs):,} PDF(s) procesados en bloques de {PROCESS_CHUNK_SIZE}, {duplicates:,} duplicado(s), {len(errors):,} incidencia(s).")
    return {"recibidos": len(uploaded_files), "pdfs": len(pdf_paths), "procesados": len(jobs), "duplicados": duplicates, "errores": len(errors)}


# ============================================================
# EXCEL DE BAJA MEMORIA
# ============================================================
def style_sheet(ws, widths=None):
    header_fill = PatternFill("solid", fgColor="111827")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    if widths:
        for idx, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = min(max(width, 10), 35)


def append_df_to_ws(ws, df: pd.DataFrame, write_header=True):
    if df is None or df.empty:
        return
    if write_header:
        ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])


def generate_excel() -> bytes:
    data = all_processed_data()
    master = build_master(data)
    nok = master[master["Resultado"] == "Nok"].copy() if not master.empty else pd.DataFrame(columns=COLUMNS)

    total_pieces = master["Pieza"].nunique() if not master.empty else 0
    nok_pieces = nok["Pieza"].nunique() if not nok.empty else 0
    ok_pieces = max(total_pieces - nok_pieces, 0)
    total_chars = len(master)
    nok_chars = len(nok)

    # Desglose solicitado: piezas rechazadas por REG 5, separando las que
    # fallaron únicamente por esta característica de las que además fallaron
    # por otra(s) característica(s).
    reg5 = "(5) 301-400 UPR"
    reg5_nok_pieces = set(nok.loc[nok["Caracteristica"].eq(reg5), "Pieza"].dropna().tolist()) if not nok.empty else set()
    reg5_only = 0
    reg5_and_others = 0
    for piece_id in reg5_nok_pieces:
        chars = set(nok.loc[nok["Pieza"].eq(piece_id), "Caracteristica"].dropna().tolist())
        if chars == {reg5}:
            reg5_only += 1
        else:
            reg5_and_others += 1

    counts = nok["Caracteristica"].value_counts().reset_index() if not nok.empty else pd.DataFrame(columns=["Caracteristica", "Cantidad_NOK"])
    counts.columns = ["Caracteristica", "Cantidad_NOK"]
    detail = []
    for char in counts["Caracteristica"].head(15):
        a = master[master["Caracteristica"] == char]
        n = nok[nok["Caracteristica"] == char]
        detail.append({
            "Caracteristica": char,
            "Cantidad Piezas NOK": n["Pieza"].nunique(),
            "Spec Min": a["Spec Min"].dropna().iloc[0] if a["Spec Min"].notna().any() else None,
            "Spec Max": a["Spec Max"].dropna().iloc[0] if a["Spec Max"].notna().any() else None,
            "Promedio Total Med.": a["Medicion"].mean(),
            "Std Total Med.": a["Medicion"].std(),
            "Promedio NOK Med.": n["Medicion"].mean() if not n.empty else None,
            "Std NOK Med.": n["Medicion"].std() if not n.empty else None,
            "Max NOK Med.": n["Medicion"].max() if not n.empty else None,
            "Min NOK Med.": n["Medicion"].min() if not n.empty else None,
        })
    detail_df = pd.DataFrame(detail)
    summary_df = pd.DataFrame({
        "Metrica": [
            "Total de piezas analizadas", "Piezas OK", "Piezas NOK", "% Rechazo",
            "FTQ", "Total de características", "Características OK", "Características NOK",
            "% Características NOK",
            'Piezas NOK por REG 5 (301-400 UPR)',
            'Piezas NOK SOLO por REG 5 (301-400 UPR)',
            'Piezas NOK por REG 5 + otras características',
            "Umbral Levas (5) 301-400 UPR", "Umbral Chatter Apoyos",
        ],
        "Valor": [
            total_pieces, ok_pieces, nok_pieces,
            (nok_pieces / total_pieces * 100) if total_pieces else 0,
            (ok_pieces / total_pieces * 100) if total_pieces else 0,
            total_chars, total_chars - nok_chars, nok_chars,
            (nok_chars / total_chars * 100) if total_chars else 0,
            len(reg5_nok_pieces), reg5_only, reg5_and_others,
            CHATTER_LEVAS_THRESHOLD, CHATTER_APOYOS_THRESHOLD,
        ],
    })

    wb = Workbook(write_only=True)
    for sheet in SHEET_ORDER:
        ws = wb.create_sheet(sheet)
        if sheet in data:
            df = data[sheet]
            append_df_to_ws(ws, df, True)
        elif sheet == "Master":
            append_df_to_ws(ws, master, True)
        elif sheet == "Nok":
            append_df_to_ws(ws, nok, True)
        elif sheet == "Analisis_Resumen":
            append_df_to_ws(ws, summary_df, True)
        elif sheet == "Analisis_Detallado_NOK":
            append_df_to_ws(ws, detail_df, True)
        elif sheet == "Errores":
            conn = db_conn()
            err_rows = conn.execute("SELECT filename,piece,error,processed_at FROM files WHERE status='ERROR' ORDER BY piece").fetchall()
            ws.append(["Nombre del archivo", "Pieza", "Error", "Fecha"])
            for row in err_rows:
                ws.append(list(row))
        else:
            ws.append(COLUMNS)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


# ============================================================
# PDF EJECUTIVO DEL DASHBOARD
# ============================================================
def _pdf_fig_bytes(fig):
    bio = io.BytesIO()
    fig.savefig(bio, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    bio.seek(0)
    return bio


def _pdf_bar(values, labels, title, xlabel="Cantidad", horizontal=False):
    fig, ax = plt.subplots(figsize=(10, 4.6))
    vals = list(values); labs = [str(x) for x in labels]
    if horizontal:
        ax.barh(labs[::-1], vals[::-1])
        ax.set_xlabel(xlabel)
    else:
        ax.bar(labs, vals)
        ax.set_ylabel(xlabel)
        ax.tick_params(axis="x", rotation=35)
    ax.set_title(title, fontweight="bold")
    ax.grid(axis="y", alpha=0.2)
    return _pdf_fig_bytes(fig)


def generate_dashboard_pdf(data: Dict[str, pd.DataFrame]) -> bytes:
    data = dict(data)
    for k in ["Apoyos", "Levas", "chatter Levas", "Chatter apoyos", "Chatter Journal Ford"]:
        data.setdefault(k, pd.DataFrame(columns=COLUMNS))
    master = build_master(data)
    if master.empty:
        raise ValueError("No hay datos procesados para generar el PDF.")
    master = master.copy()
    master["Medicion"] = pd.to_numeric(master["Medicion"], errors="coerce")
    nok = master[master["Resultado"] == "Nok"].copy()
    total = int(master["Pieza"].nunique())
    nok_pieces = int(nok["Pieza"].nunique())
    ok_pieces = max(total - nok_pieces, 0)
    ftq = ok_pieces / total * 100 if total else 0
    rejection = nok_pieces / total * 100 if total else 0

    reg5 = "(5) 301-400 UPR"
    reg5_nok = nok[nok["Caracteristica"].eq(reg5)].copy()
    reg5_piece_ids = set(reg5_nok["Pieza"].dropna().tolist())
    reg5_only = 0; reg5_others = 0
    for pid in reg5_piece_ids:
        chars = set(nok.loc[nok["Pieza"].eq(pid), "Caracteristica"].dropna().tolist())
        if chars == {reg5}: reg5_only += 1
        else: reg5_others += 1

    # Separación explícita de REG 5 por Base Circle y Lift Area.
    reg5_area = (reg5_nok[reg5_nok["Area"].isin(["Base Circle", "Lift Area"])]
                 .groupby("Area")["Pieza"].nunique()
                 .reindex(["Base Circle", "Lift Area"], fill_value=0))
    reg5_area_rows = [
        ["REG 5 — Base Circle", int(reg5_area.get("Base Circle", 0))],
        ["REG 5 — Lift Area", int(reg5_area.get("Lift Area", 0))],
        ["REG 5 — Total de piezas NOK", len(reg5_piece_ids)],
        ["REG 5 — Solo REG 5", reg5_only],
        ["REG 5 — REG 5 + otras", reg5_others],
    ]

    top = nok.groupby("Caracteristica").size().sort_values(ascending=False).head(12)
    chatter_levas = data["chatter Levas"].copy()
    chatter_apoyos = data["Chatter apoyos"].copy()
    cln = chatter_levas[chatter_levas["Resultado"] == "Nok"] if not chatter_levas.empty else chatter_levas
    can = chatter_apoyos[chatter_apoyos["Resultado"] == "Nok"] if not chatter_apoyos.empty else chatter_apoyos
    chatter_levas_by = cln.groupby("Leva").size().sort_values(ascending=False) if not cln.empty else pd.Series(dtype=int)
    chatter_apoyos_by = can.groupby("Apoyo").size().sort_values(ascending=False) if not can.empty else pd.Series(dtype=int)

    pdf = io.BytesIO()
    doc = SimpleDocTemplate(pdf, pagesize=landscape(A4), rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("Title2", parent=styles["Title"], alignment=TA_CENTER, fontSize=20, leading=24, spaceAfter=10)
    h = ParagraphStyle("H2x", parent=styles["Heading2"], fontSize=14, leading=17, spaceBefore=6, spaceAfter=8)
    small = ParagraphStyle("smallx", parent=styles["BodyText"], fontSize=8, leading=10)
    story = []
    story.append(Paragraph("Reporte NoK — Dashboard Ejecutivo de Calidad", title))
    story.append(Paragraph(f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Versión {APP_VERSION}", small))
    story.append(Spacer(1, 10))

    metrics = [
        ["Piezas analizadas", f"{total:,}", "Piezas OK", f"{ok_pieces:,}", "Piezas NOK", f"{nok_pieces:,}"],
        ["FTQ", f"{ftq:.2f}%", "Rechazo", f"{rejection:.2f}%", "Mediciones NOK", f"{len(nok):,}"],
    ]
    mt = Table(metrics, colWidths=[110, 70, 90, 70, 95, 70])
    mt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F3F4F6")),("GRID",(0,0),(-1,-1),0.4,colors.grey),("FONTNAME",(0,0),(-1,-1),"Helvetica"),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("ALIGN",(1,0),(1,-1),"CENTER"),("ALIGN",(3,0),(3,-1),"CENTER"),("ALIGN",(5,0),(5,-1),"CENTER")]))
    story.append(mt); story.append(Spacer(1, 10))

    okfig = _pdf_bar([ok_pieces, nok_pieces], ["OK", "NOK"], "Piezas OK vs NOK", horizontal=False)
    story.append(RLImage(okfig, width=360, height=165))
    story.append(Spacer(1, 6))
    story.append(Paragraph("REG 5 — (5) 301-400 UPR", h))
    rt = Table([["Indicador", "Piezas NOK"]] + reg5_area_rows, colWidths=[260, 100])
    rt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#111827")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.4,colors.grey),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("ALIGN",(1,1),(1,-1),"CENTER")]))
    story.append(rt)
    # Top piezas NOK para trazabilidad ejecutiva. El detalle completo queda en Excel.
    piece_counts = nok.groupby("Pieza").size().sort_values(ascending=False).head(12) if not nok.empty else pd.Series(dtype=int)
    story.append(Spacer(1, 8))
    story.append(Paragraph("Top piezas con mayor número de mediciones NOK", h))
    if not piece_counts.empty:
        pt = [["Pieza", "Mediciones NOK"]] + [[str(k), int(v)] for k, v in piece_counts.items()]
        ptable = Table(pt, colWidths=[180, 110])
        ptable.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#111827")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.35,colors.grey),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("ALIGN",(1,1),(1,-1),"CENTER")]))
        story.append(ptable)
    if not reg5_area.empty:
        figb = _pdf_bar(reg5_area.values, reg5_area.index, "REG 5 NOK por área", horizontal=False)
        story.append(Spacer(1, 8)); story.append(RLImage(figb, width=430, height=190))
    story.append(PageBreak())

    story.append(Paragraph("Características NOK", h))
    if not top.empty:
        f = _pdf_bar(top.values, top.index, "Top características NOK", horizontal=True)
        story.append(RLImage(f, width=650, height=285))
    else:
        story.append(Paragraph("No se encontraron características NOK.", small))
    story.append(Spacer(1, 8))
    story.append(Paragraph("Regla activa: REG 5 en Levas = umbral 0.0001; el reporte conserva Base Circle y Lift Area por separado.", small))
    story.append(PageBreak())

    story.append(Paragraph("Chatter Levas", h))
    if not chatter_levas_by.empty:
        f = _pdf_bar(chatter_levas_by.values, chatter_levas_by.index, "Chatter Levas — NOK por Leva", horizontal=False)
        story.append(RLImage(f, width=650, height=285))
    else: story.append(Paragraph("No hay Chatter Levas NOK.", small))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Chatter Apoyos", h))
    if not chatter_apoyos_by.empty:
        f = _pdf_bar(chatter_apoyos_by.values, chatter_apoyos_by.index, "Chatter Apoyos — NOK por Apoyo", horizontal=False)
        story.append(RLImage(f, width=650, height=285))
    else: story.append(Paragraph("No hay Chatter Apoyos NOK.", small))
    story.append(Spacer(1, 8))
    story.append(Paragraph("Este PDF es un resumen ejecutivo; el Excel conserva el detalle completo por pieza, característica, leva y apoyo.", small))
    doc.build(story)
    pdf.seek(0)
    return pdf.getvalue()


# ============================================================
# DASHBOARD
# ============================================================
def dashboard(data: Dict[str, pd.DataFrame]):
    """Dashboard ejecutivo. Los bloques internos de 150 son solo de procesamiento;
    aquí se consolida y muestra el resultado individual de todas las piezas."""
    for _k in ["Apoyos", "Levas", "chatter Levas", "Chatter apoyos", "Chatter Journal Ford"]:
        data.setdefault(_k, pd.DataFrame(columns=COLUMNS))
    master = build_master(data)
    if master.empty:
        st.info("Procesa al menos un lote para activar el dashboard.")
        return

    master = master.copy()
    master["Medicion"] = pd.to_numeric(master["Medicion"], errors="coerce")
    nok = master[master["Resultado"].eq("Nok")].copy()
    total = int(master["Pieza"].nunique())
    nok_pieces = int(nok["Pieza"].nunique())
    ok_pieces = max(total - nok_pieces, 0)
    ftq = ok_pieces / total * 100 if total else 0
    rejection = nok_pieces / total * 100 if total else 0

    # ---------------- FILTROS ----------------
    with st.sidebar:
        st.markdown("### 🎛️ Filtros del análisis")
        areas = sorted(master["Area"].dropna().astype(str).unique().tolist())
        selected_areas = st.multiselect("Área", areas, default=areas, key="dash_areas")
        results = st.multiselect("Resultado", ["Ok", "Nok"], default=["Ok", "Nok"], key="dash_results")
        chars = sorted(master["Caracteristica"].dropna().astype(str).unique().tolist())
        selected_chars = st.multiselect("Característica", chars, default=chars, key="dash_chars")

    view = master[master["Area"].isin(selected_areas) & master["Resultado"].isin(results)].copy()
    if selected_chars:
        view = view[view["Caracteristica"].isin(selected_chars)]
    view_nok = view[view["Resultado"].eq("Nok")]

    # ---------------- ENCABEZADO EJECUTIVO ----------------
    st.markdown("## 📊 Executive Quality Dashboard")
    st.caption(f"Reporte consolidado · {total:,} piezas · Los bloques de 150 son internos y no fragmentan el resultado final")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Piezas analizadas", f"{total:,}")
    c2.metric("Piezas OK", f"{ok_pieces:,}")
    c3.metric("Piezas NOK", f"{nok_pieces:,}")
    c4.metric("FTQ", f"{ftq:.2f}%")
    c5.metric("Rechazo", f"{rejection:.2f}%")

    # ---------------- REG 5 ----------------
    reg5 = "(5) 301-400 UPR"
    reg5_nok = nok[nok["Caracteristica"].eq(reg5)].copy()
    reg5_piece_ids = set(reg5_nok["Pieza"].dropna().tolist())
    reg5_only_ids, reg5_other_ids = set(), set()
    for pid in reg5_piece_ids:
        chars_piece = set(nok.loc[nok["Pieza"].eq(pid), "Caracteristica"].dropna().astype(str))
        if chars_piece == {reg5}:
            reg5_only_ids.add(pid)
        else:
            reg5_other_ids.add(pid)
    reg5_area = (reg5_nok[reg5_nok["Area"].isin(["Base Circle", "Lift Area"])]
                 .groupby("Area")["Pieza"].nunique()
                 .reindex(["Base Circle", "Lift Area"], fill_value=0))

    st.markdown("### 🎯 REG 5 — (5) 301-400 UPR")
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("REG 5 NOK", f"{len(reg5_piece_ids):,}")
    r2.metric("Base Circle NOK", f"{int(reg5_area.get('Base Circle', 0)):,}")
    r3.metric("Lift Area NOK", f"{int(reg5_area.get('Lift Area', 0)):,}")
    r4.metric("REG 5 + otras", f"{len(reg5_other_ids):,}")
    st.caption(f"Solo REG 5: {len(reg5_only_ids):,} piezas · Umbral Levas: ≥ 0.0001")

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "🏁 Ejecutivo", "🎯 REG 5", "🟦 Levas", "🟩 Apoyos", "🔥 Chatter", "📋 Piezas"
    ])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            counts = pd.DataFrame({"Resultado": ["OK", "NOK"], "Cantidad": [ok_pieces, nok_pieces]})
            fig = px.pie(counts, names="Resultado", values="Cantidad", hole=0.60, title="Distribución de piezas")
            fig.update_layout(height=390, margin=dict(l=20,r=20,t=55,b=20), legend_title_text="")
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            by_char = view_nok.groupby("Caracteristica").size().reset_index(name="NOK").sort_values("NOK", ascending=False).head(15)
            fig = px.bar(by_char, x="NOK", y="Caracteristica", orientation="h", title="Top 15 características NOK")
            fig.update_layout(height=390, yaxis={"categoryorder":"total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        by_area = view_nok.groupby("Area").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
        if not by_area.empty:
            fig = px.bar(by_area, x="Area", y="NOK", title="NOK por área")
            st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.markdown("#### Separación obligatoria de REG 5 por área")
        reg5_tbl = pd.DataFrame({
            "Área": ["Base Circle", "Lift Area"],
            "Piezas NOK": [int(reg5_area.get("Base Circle",0)), int(reg5_area.get("Lift Area",0))],
            "Mediciones NOK": [int(((reg5_nok["Area"]=="Base Circle")).sum()), int(((reg5_nok["Area"]=="Lift Area")).sum())],
        })
        st.dataframe(reg5_tbl, hide_index=True, use_container_width=True)
        if reg5_area.sum() > 0:
            fig = px.bar(reg5_tbl, x="Área", y="Piezas NOK", title="REG 5 — piezas NOK: Base Circle vs Lift Area", text="Piezas NOK")
            st.plotly_chart(fig, use_container_width=True)
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Piezas NOK SOLO por REG 5", f"{len(reg5_only_ids):,}")
        with col2:
            st.metric("Piezas NOK REG 5 + otras características", f"{len(reg5_other_ids):,}")
        if not reg5_nok.empty:
            st.dataframe(reg5_nok.sort_values(["Area","Pieza"]), hide_index=True, use_container_width=True, height=350)

    with tab3:
        levas = view[view["Area"].isin(["Levas", "Base Circle", "Lift Area"])].copy()
        if levas.empty:
            st.info("No hay datos de Levas en los filtros actuales.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                n = levas[levas.Resultado=="Nok"].groupby("Leva").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                st.plotly_chart(px.bar(n, x="Leva", y="NOK", title="NOK por Leva"), use_container_width=True)
            with col2:
                n = levas[levas.Resultado=="Nok"].groupby("Caracteristica").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                fig = px.bar(n, x="NOK", y="Caracteristica", orientation="h", title="NOK por característica de Levas")
                fig.update_layout(yaxis={"categoryorder":"total ascending"})
                st.plotly_chart(fig, use_container_width=True)
            heat = levas[levas.Resultado=="Nok"].pivot_table(index="Caracteristica", columns="Leva", values="Pieza", aggfunc="nunique", fill_value=0)
            if not heat.empty:
                st.plotly_chart(px.imshow(heat, aspect="auto", title="Mapa de calor — piezas NOK por característica × Leva"), use_container_width=True)

    with tab4:
        apoyos = view[view["Area"].eq("Apoyos")].copy()
        if apoyos.empty:
            st.info("No hay datos de Apoyos en los filtros actuales.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                n = apoyos[apoyos.Resultado=="Nok"].groupby("Apoyo").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                st.plotly_chart(px.bar(n, x="Apoyo", y="NOK", title="NOK por Apoyo"), use_container_width=True)
            with col2:
                n = apoyos[apoyos.Resultado=="Nok"].groupby("Caracteristica").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                fig = px.bar(n, x="NOK", y="Caracteristica", orientation="h", title="NOK por característica de Apoyos")
                fig.update_layout(yaxis={"categoryorder":"total ascending"})
                st.plotly_chart(fig, use_container_width=True)
            heat = apoyos[apoyos.Resultado=="Nok"].pivot_table(index="Caracteristica", columns="Apoyo", values="Pieza", aggfunc="nunique", fill_value=0)
            if not heat.empty:
                st.plotly_chart(px.imshow(heat, aspect="auto", title="Mapa de calor — piezas NOK por característica × Apoyo"), use_container_width=True)

    with tab5:
        st.markdown("### 🔥 Chatter — análisis separado")
        chatter_levas = data.get("chatter Levas", pd.DataFrame(columns=COLUMNS)).copy()
        chatter_apoyos = data.get("Chatter apoyos", pd.DataFrame(columns=COLUMNS)).copy()
        c1, c2 = st.columns(2)
        with c1:
            if chatter_levas.empty: st.info("No hay datos de Chatter Levas.")
            else:
                n = chatter_levas[chatter_levas.Resultado=="Nok"].groupby("Leva").size().reset_index(name="NOK")
                st.plotly_chart(px.bar(n, x="Leva", y="NOK", title="Chatter Levas — NOK por Leva"), use_container_width=True)
        with c2:
            if chatter_apoyos.empty: st.info("No hay datos de Chatter Apoyos.")
            else:
                n = chatter_apoyos[chatter_apoyos.Resultado=="Nok"].groupby("Apoyo").size().reset_index(name="NOK")
                st.plotly_chart(px.bar(n, x="Apoyo", y="NOK", title="Chatter Apoyos — NOK por Apoyo"), use_container_width=True)
        c1, c2 = st.columns(2)
        with c1:
            if not chatter_levas.empty:
                n = chatter_levas[chatter_levas.Resultado=="Nok"].groupby(["Caracteristica","Area"]).size().reset_index(name="NOK")
                st.plotly_chart(px.bar(n, x="Caracteristica", y="NOK", color="Area", barmode="group", title="Chatter Levas — Base Circle vs Lift Area"), use_container_width=True)
        with c2:
            if not chatter_apoyos.empty:
                n = chatter_apoyos[chatter_apoyos.Resultado=="Nok"].groupby("Caracteristica").size().reset_index(name="NOK")
                fig = px.bar(n, x="NOK", y="Caracteristica", orientation="h", title="Chatter Apoyos — NOK por característica")
                fig.update_layout(yaxis={"categoryorder":"total ascending"})
                st.plotly_chart(fig, use_container_width=True)
        frames=[]
        if not chatter_levas.empty:
            x=chatter_levas.copy(); x["Tipo"]="Leva"; frames.append(x)
        if not chatter_apoyos.empty:
            x=chatter_apoyos.copy(); x["Tipo"]="Apoyo"; frames.append(x)
        if frames:
            cf=pd.concat(frames, ignore_index=True)
            chars_chatter=sorted(cf["Caracteristica"].dropna().astype(str).unique())
            selected=st.selectbox("Característica de Chatter", chars_chatter, key="v95_chatter_char")
            cd=cf[cf["Caracteristica"].astype(str).eq(selected)]
            if not cd.empty:
                fig=px.histogram(cd,x="Medicion",color="Tipo",marginal="box",nbins=40,title=f"Distribución de medición — {selected}")
                fig.add_vline(x=CHATTER_LEVAS_THRESHOLD,line_dash="dash",annotation_text="Levas 0.0001")
                fig.add_vline(x=CHATTER_APOYOS_THRESHOLD,line_dash="dot",annotation_text="Apoyos 0.00008")
                st.plotly_chart(fig,use_container_width=True)

    with tab6:
        st.markdown("### 📋 Resultado individual por pieza")
        st.caption("Esta tabla está consolidada de todos los bloques internos. No representa lotes de 150.")
        piece_rows=[]
        for pid, g in master.groupby("Pieza", sort=True):
            bad=g[g["Resultado"].eq("Nok")]
            bad_chars=sorted(set(bad["Caracteristica"].dropna().astype(str)))
            piece_rows.append({
                "Pieza": pid,
                "Resultado": "NOK" if not bad.empty else "OK",
                "NOK mediciones": len(bad),
                "Características NOK": ", ".join(bad_chars) if bad_chars else "—",
                "REG 5": "NOK" if reg5 in bad_chars else "OK",
                "REG 5 área": ", ".join(sorted(set(bad.loc[bad["Caracteristica"].eq(reg5),"Area"].dropna().astype(str)))) if reg5 in bad_chars else "—",
                "Archivo PDF": str(g["Nombre del archivo"].iloc[0]),
            })
        pieces_df=pd.DataFrame(piece_rows)
        piece_filter=st.multiselect("Mostrar", ["OK","NOK"], default=["OK","NOK"], key="piece_result_filter")
        search=st.text_input("🔎 Buscar pieza o archivo PDF", key="piece_search").strip().lower()
        pieces_view=pieces_df[pieces_df["Resultado"].isin(piece_filter)].copy()
        if search:
            mask=pieces_view.apply(lambda r: search in str(r["Pieza"]).lower() or search in str(r["Archivo PDF"]).lower(), axis=1)
            pieces_view=pieces_view[mask]
        st.metric("Piezas mostradas", f"{len(pieces_view):,}")
        st.dataframe(pieces_view, hide_index=True, use_container_width=True, height=520)

    st.markdown("### 🔍 Registros NOK individuales")
    if not nok.empty:
        st.dataframe(nok.sort_values(["Pieza", "Caracteristica"]), hide_index=True, use_container_width=True, height=360)


# ============================================================
# UI PRINCIPAL
# ============================================================
with st.sidebar:
    st.markdown("## ⚙️ Configuración")
    st.metric("Máximo de PDFs por ejecución", f"{MAX_INPUT_PDFS:,}")
    st.metric("Bloque interno de procesamiento", f"{PROCESS_CHUNK_SIZE}")
    workers = st.slider("Workers paralelos", min_value=1, max_value=MAX_WORKERS, value=8)
    st.caption("Puedes seleccionar hasta 1,500 PDFs de una vez. El motor los procesa en bloques de 150, usa 8 workers por defecto y reduce las escrituras a disco para acelerar el proceso.")
    st.divider()
    st.markdown("### Reglas activas")
    st.code("Levas (5) 301-400 UPR ≥ 0.0001 → NOK\nChatter Apoyos → 0.00008", language="text")
    if st.button("🗑️ Nueva sesión / limpiar resultados", use_container_width=True):
        reset_workspace()

st.markdown("### 1️⃣ Cargar archivos")
st.info("Modo optimizado: puedes seleccionar hasta 1,500 PDFs en una sola ejecución. La aplicación los procesa internamente en bloques de 150, libera memoria entre bloques y genera un único resultado acumulado. También acepta uno o varios ZIP y mezcla PDF + ZIP.")

uploaded_files = st.file_uploader(
    "Selecciona PDFs y/o ZIPs",
    type=["pdf", "zip"],
    accept_multiple_files=True,
    help="Puedes seleccionar hasta 1,500 PDFs. Internamente se procesan en bloques de 150 para reducir el uso de RAM. Los ZIP se extraen y procesan en disco temporal.",
)

col_a, col_b = st.columns([1, 1])
with col_a:
    if st.button("🚀 Procesar lote", type="primary", disabled=not uploaded_files, use_container_width=True):
        st.session_state.last_summary = process_uploaded_batch(uploaded_files, int(workers))
with col_b:
    conn = db_conn()
    processed_count = conn.execute("SELECT COUNT(*) FROM files WHERE status='OK'").fetchone()[0]
    duplicate_count = 0  # se refleja por lote; los hashes previos quedan registrados
    st.metric("PDFs procesados acumulados", f"{processed_count:,}")

# Resumen de incidencias
conn = db_conn()
error_count = conn.execute("SELECT COUNT(*) FROM files WHERE status='ERROR'").fetchone()[0]
processed_count = conn.execute("SELECT COUNT(*) FROM files WHERE status='OK'").fetchone()[0]
if processed_count:
    st.success(f"Repositorio de sesión: {processed_count:,} PDF(s) procesados correctamente | {error_count:,} con error.")

st.markdown("### 2️⃣ Dashboard")
data_for_dashboard = all_processed_data()
dashboard(data_for_dashboard)

st.markdown("### 3️⃣ Exportación")
col1, col2, col3 = st.columns(3)
with col1:
    if st.button("📗 Generar Excel final", use_container_width=True, disabled=processed_count == 0):
        with st.spinner("Generando Excel de baja memoria…"):
            excel_bytes = generate_excel()
        st.session_state.excel_bytes = excel_bytes
        st.success("Excel generado correctamente.")
with col2:
    if "excel_bytes" in st.session_state:
        st.download_button(
            "⬇️ Descargar Reporte_NoK.xlsx",
            data=st.session_state.excel_bytes,
            file_name=f"Reporte_NoK_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
with col3:
    if st.button("📄 Generar PDF del Dashboard", use_container_width=True, disabled=processed_count == 0):
        with st.spinner("Generando PDF ejecutivo…"):
            st.session_state.dashboard_pdf = generate_dashboard_pdf(data_for_dashboard)
        st.success("PDF ejecutivo generado.")
    if "dashboard_pdf" in st.session_state:
        st.download_button(
            "⬇️ Descargar Dashboard PDF",
            data=st.session_state.dashboard_pdf,
            file_name=f"Reporte_NoK_Dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

if error_count:
    with st.expander(f"⚠️ Ver errores ({error_count})"):
        rows = conn.execute("SELECT filename,piece,error,processed_at FROM files WHERE status='ERROR' ORDER BY piece").fetchall()
        st.dataframe(pd.DataFrame(rows, columns=["Archivo","Pieza","Error","Fecha"]), use_container_width=True)

st.caption(f"Reporte NoK {APP_VERSION} · procesamiento por lotes · SHA-256 · {datetime.now().strftime('%Y-%m-%d %H:%M')}")

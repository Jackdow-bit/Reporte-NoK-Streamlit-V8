# -*- coding: utf-8 -*-
"""
Reporte NoK Streamlit V9.1 - optimizado para Streamlit Community Cloud

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


# ============================================================
# CONFIGURACIÓN
# ============================================================
APP_VERSION = "V10.0"
MAX_ZIP_DEPTH = 8
MAX_INPUT_PDFS = 1500
MAX_WORKERS = 6

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


def register_file(conn, sha256, filename, piece, status="OK", error=""):
    conn.execute(
        "INSERT OR REPLACE INTO files(sha256,filename,piece,status,error,processed_at) VALUES(?,?,?,?,?,?)",
        (sha256, filename, piece, status, error, datetime.now().isoformat(timespec="seconds")),
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


def save_individual_result(pdf_dir: Path, frames: Dict[str, pd.DataFrame]):
    pdf_dir.mkdir(parents=True, exist_ok=True)
    for sheet, df in frames.items():
        if df is None or df.empty:
            continue
        safe_sheet = re.sub(r"[^A-Za-z0-9_-]+", "_", sheet)
        df.to_csv(pdf_dir / f"{safe_sheet}.csv.gz", index=False, compression="gzip")


def read_all_sheet(sheet: str) -> pd.DataFrame:
    pattern = f"{re.sub(r'[^A-Za-z0-9_-]+','_',sheet)}.csv.gz"
    files = sorted((WORKSPACE / "batches").glob(f"**/{pattern}"))
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
        st.write(f"Recibidos: {len(uploaded_files):,} archivo(s)")
        pdf_paths, input_errors = materialize_uploaded_sources(uploaded_files, source_dir)
        st.write(f"PDFs encontrados en PDF/ZIP: {len(pdf_paths):,}")
        if input_errors:
            st.warning(f"{len(input_errors):,} entrada(s) no se pudieron preparar.")

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
    progress = st.progress(0, text="Procesando PDFs individualmente…")
    started = time.perf_counter()
    completed = 0
    jobs = [(p, name, num, digest) for p, name, num, digest in unique_jobs]

    # Cada PDF conserva su resultado individual. No se concatena ni se descarta
    # información por bloques de 150. Los resultados de cada PDF se escriben
    # inmediatamente en su propia carpeta persistente del workspace.
    with st.status("Procesamiento individual", expanded=False) as status:
        with ThreadPoolExecutor(max_workers=min(max(1, workers), max(1, len(jobs) or 1))) as executor:
            future_map = {
                executor.submit(process_pdf_job, (p, name, num)): (p, name, num, digest)
                for p, name, num, digest in jobs
            }
            for future in as_completed(future_map):
                p, name, num, digest = future_map[future]
                try:
                    out = future.result()
                    if out["ok"]:
                        individual_dir = batch_dir / f"Pieza_{num:05d}_{digest[:8]}"
                        frames = dict(out["data"])
                        # Chatter Journal Ford se conserva como una hoja individual
                        # derivada del mismo PDF, sin perder el origen.
                        ford = make_ford_df(frames.get("Chatter apoyos", pd.DataFrame(columns=COLUMNS)))
                        if not ford.empty:
                            frames["Chatter Journal Ford"] = ford
                        save_individual_result(individual_dir, frames)
                        register_file(conn, digest, name, num, "OK", "")
                    else:
                        register_file(conn, digest, name, num, "ERROR", out["error"])
                        errors.append(f"{name}: {out['error']}")
                except Exception as e:
                    err = f"{type(e).__name__}: {e}"
                    register_file(conn, digest, name, num, "ERROR", err)
                    errors.append(f"{name}: {err}")

                completed += 1
                elapsed = max(time.perf_counter() - started, 0.001)
                rate = completed / elapsed
                eta = (len(jobs) - completed) / rate if rate else 0
                progress.progress(
                    completed / max(len(jobs), 1),
                    text=f"{completed:,}/{len(jobs):,} PDFs | {rate:.1f} PDF/s | ETA {eta/60:.1f} min"
                )

        status.update(label=f"Procesamiento terminado: {completed:,} PDF(s)", state="complete")

    (batch_dir / "errors.json").write_text(json.dumps(errors, ensure_ascii=False, indent=2), encoding="utf-8")
    (batch_dir / "meta.json").write_text(json.dumps({
        "batch_id": batch_id,
        "files": len(uploaded_files),
        "pdfs": len(pdf_paths),
        "processed": len(jobs),
        "duplicates": duplicates,
        "errors": len(errors),
        "elapsed_s": time.perf_counter() - started,
        "input_limit": MAX_INPUT_PDFS,
        "processing_mode": "individual_pdf",
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    shutil.rmtree(source_dir, ignore_errors=True)

    st.success(
        f"Carga terminada: {len(jobs):,} PDF(s) procesados individualmente, "
        f"{duplicates:,} duplicado(s), {len(errors):,} incidencia(s)."
    )
    return {
        "recibidos": len(uploaded_files),
        "pdfs": len(pdf_paths),
        "procesados": len(jobs),
        "duplicados": duplicates,
        "errores": len(errors),
    }


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
            "% Características NOK", "Umbral Levas (5) 301-400 UPR", "Umbral Chatter Apoyos",
        ],
        "Valor": [
            total_pieces, ok_pieces, nok_pieces,
            (nok_pieces / total_pieces * 100) if total_pieces else 0,
            (ok_pieces / total_pieces * 100) if total_pieces else 0,
            total_chars, total_chars - nok_chars, nok_chars,
            (nok_chars / total_chars * 100) if total_chars else 0,
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
# DASHBOARD
# ============================================================
def dashboard(data: Dict[str, pd.DataFrame]):
    for _k in ["Apoyos", "Levas", "chatter Levas", "Chatter apoyos", "Chatter Journal Ford"]:
        data.setdefault(_k, pd.DataFrame(columns=COLUMNS))
    master = build_master(data)
    if master.empty:
        st.info("Procesa al menos un lote para activar el dashboard.")
        return

    master["Medicion"] = pd.to_numeric(master["Medicion"], errors="coerce")
    nok = master[master["Resultado"] == "Nok"].copy()
    total = master["Pieza"].nunique()
    nok_pieces = nok["Pieza"].nunique()
    ftq = (total - nok_pieces) / total * 100 if total else 0
    rejection = nok_pieces / total * 100 if total else 0

    # Filtros
    with st.sidebar:
        st.markdown("### 🎛️ Filtros del análisis")
        areas = sorted(master["Area"].dropna().astype(str).unique().tolist())
        selected_areas = st.multiselect("Área", areas, default=areas)
        chars = sorted(master["Caracteristica"].dropna().astype(str).unique().tolist())
        selected_chars = st.multiselect("Característica", chars, default=chars[:20] if len(chars) > 20 else chars)
        results = st.multiselect("Resultado", ["Ok", "Nok"], default=["Ok", "Nok"])

    view = master[master["Area"].isin(selected_areas) & master["Resultado"].isin(results)].copy()
    if selected_chars:
        view = view[view["Caracteristica"].isin(selected_chars)]
    view_nok = view[view["Resultado"] == "Nok"]

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Piezas", f"{total:,}")
    c2.metric("Piezas NOK", f"{nok_pieces:,}")
    c3.metric("FTQ", f"{ftq:.2f}%")
    c4.metric("Rechazo", f"{rejection:.2f}%")
    c5.metric("NOK mediciones", f"{len(nok):,}")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["🏁 Ejecutivo", "🟦 Levas", "🟩 Apoyos", "🔥 Chatter", "🔎 Características"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            counts = pd.DataFrame({"Resultado": ["OK", "NOK"], "Cantidad": [len(view[view.Resultado=="Ok"]), len(view[view.Resultado=="Nok"])]})
            fig = px.pie(counts, names="Resultado", values="Cantidad", hole=0.58, title="Distribución OK / NOK")
            fig.update_layout(height=390, legend_title_text="")
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            by_char = view_nok.groupby("Caracteristica").size().reset_index(name="NOK").sort_values("NOK", ascending=False).head(15)
            fig = px.bar(by_char, x="NOK", y="Caracteristica", orientation="h", title="Top características NOK")
            fig.update_layout(height=390, yaxis={"categoryorder":"total ascending"})
            st.plotly_chart(fig, use_container_width=True)

        by_area = view_nok.groupby("Area").size().reset_index(name="NOK")
        fig = px.bar(by_area, x="Area", y="NOK", title="NOK por área")
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        levas = view[view["Area"].isin(["Levas", "Base Circle", "Lift Area"])].copy()
        if levas.empty:
            st.info("No hay datos de Levas en los filtros actuales.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                n = levas[levas.Resultado=="Nok"].groupby("Leva").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                fig = px.bar(n, x="Leva", y="NOK", title="NOK por Leva")
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                n = levas[levas.Resultado=="Nok"].groupby("Caracteristica").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                fig = px.bar(n, x="NOK", y="Caracteristica", orientation="h", title="NOK por característica de Levas")
                fig.update_layout(yaxis={"categoryorder":"total ascending"})
                st.plotly_chart(fig, use_container_width=True)

            if "Leva" in levas.columns:
                heat = levas[levas.Resultado=="Nok"].pivot_table(index="Caracteristica", columns="Leva", values="Pieza", aggfunc="nunique", fill_value=0)
                if not heat.empty:
                    fig = px.imshow(heat, aspect="auto", title="Mapa de calor: piezas NOK por característica × Leva", labels=dict(x="Leva", y="Característica", color="Piezas NOK"))
                    st.plotly_chart(fig, use_container_width=True)

            chatter = data.get("chatter Levas", pd.DataFrame(columns=COLUMNS))
            if not chatter.empty:
                chatter_nok = chatter[chatter.Resultado=="Nok"]
                n = chatter_nok.groupby(["Area","Caracteristica"]).size().reset_index(name="NOK")
                fig = px.bar(n, x="Caracteristica", y="NOK", color="Area", barmode="group", title="Chatter Levas: NOK por característica y área")
                st.plotly_chart(fig, use_container_width=True)

    with tab3:
        apoyos = view[view["Area"] == "Apoyos"].copy()
        if apoyos.empty:
            st.info("No hay datos de Apoyos en los filtros actuales.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                n = apoyos[apoyos.Resultado=="Nok"].groupby("Apoyo").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                fig = px.bar(n, x="Apoyo", y="NOK", title="NOK por Apoyo")
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                n = apoyos[apoyos.Resultado=="Nok"].groupby("Caracteristica").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                fig = px.bar(n, x="NOK", y="Caracteristica", orientation="h", title="NOK por característica de Apoyos")
                fig.update_layout(yaxis={"categoryorder":"total ascending"})
                st.plotly_chart(fig, use_container_width=True)

            heat = apoyos[apoyos.Resultado=="Nok"].pivot_table(index="Caracteristica", columns="Apoyo", values="Pieza", aggfunc="nunique", fill_value=0)
            if not heat.empty:
                fig = px.imshow(heat, aspect="auto", title="Mapa de calor: piezas NOK por característica × Apoyo", labels=dict(x="Apoyo", y="Característica", color="Piezas NOK"))
                st.plotly_chart(fig, use_container_width=True)

            chatter = data.get("Chatter apoyos", pd.DataFrame(columns=COLUMNS))
            if not chatter.empty:
                chatter_nok = chatter[chatter.Resultado=="Nok"]
                n = chatter_nok.groupby("Caracteristica").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                fig = px.bar(n, x="NOK", y="Caracteristica", orientation="h", title="Chatter Apoyos: NOK por característica")
                fig.update_layout(yaxis={"categoryorder":"total ascending"})
                st.plotly_chart(fig, use_container_width=True)

    with tab4:
        st.markdown("### 🔥 Análisis específico de Chatter")
        chatter_levas = data.get("chatter Levas", pd.DataFrame(columns=COLUMNS)).copy()
        chatter_apoyos = data.get("Chatter apoyos", pd.DataFrame(columns=COLUMNS)).copy()
        chatter_levas["Medicion"] = pd.to_numeric(chatter_levas.get("Medicion", pd.Series(dtype=float)), errors="coerce") if not chatter_levas.empty else chatter_levas.get("Medicion", pd.Series(dtype=float))
        chatter_apoyos["Medicion"] = pd.to_numeric(chatter_apoyos.get("Medicion", pd.Series(dtype=float)), errors="coerce") if not chatter_apoyos.empty else chatter_apoyos.get("Medicion", pd.Series(dtype=float))

        c1, c2 = st.columns(2)
        with c1:
            if chatter_levas.empty:
                st.info("No hay datos de Chatter Levas.")
            else:
                n = chatter_levas[chatter_levas["Resultado"] == "Nok"].groupby("Leva").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                fig = px.bar(n, x="Leva", y="NOK", title="Chatter Levas — NOK por Leva")
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            if chatter_apoyos.empty:
                st.info("No hay datos de Chatter Apoyos.")
            else:
                n = chatter_apoyos[chatter_apoyos["Resultado"] == "Nok"].groupby("Apoyo").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                fig = px.bar(n, x="Apoyo", y="NOK", title="Chatter Apoyos — NOK por Apoyo")
                st.plotly_chart(fig, use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            if not chatter_levas.empty:
                n = chatter_levas[chatter_levas["Resultado"] == "Nok"].groupby(["Caracteristica", "Area"]).size().reset_index(name="NOK")
                fig = px.bar(n, x="Caracteristica", y="NOK", color="Area", barmode="group", title="Chatter Levas — NOK por característica y área")
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            if not chatter_apoyos.empty:
                n = chatter_apoyos[chatter_apoyos["Resultado"] == "Nok"].groupby("Caracteristica").size().reset_index(name="NOK").sort_values("NOK", ascending=False)
                fig = px.bar(n, x="NOK", y="Caracteristica", orientation="h", title="Chatter Apoyos — NOK por característica")
                fig.update_layout(yaxis={"categoryorder": "total ascending"})
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Distribución de medición de Chatter")
        chatter_frames = []
        if not chatter_levas.empty:
            x = chatter_levas.copy(); x["Tipo"] = "Leva"; chatter_frames.append(x)
        if not chatter_apoyos.empty:
            x = chatter_apoyos.copy(); x["Tipo"] = "Apoyo"; chatter_frames.append(x)
        if chatter_frames:
            cf = pd.concat(chatter_frames, ignore_index=True)
            chars_chatter = sorted(cf["Caracteristica"].dropna().astype(str).unique())
            selected_chatter_char = st.selectbox("Característica de Chatter", chars_chatter, key="v91_chatter_char")
            cd = cf[cf["Caracteristica"].astype(str) == selected_chatter_char].copy()
            if not cd.empty:
                fig = px.histogram(cd, x="Medicion", color="Tipo", marginal="box", nbins=40, title=f"Distribución — {selected_chatter_char}")
                fig.add_vline(x=CHATTER_LEVAS_THRESHOLD, line_dash="dash", annotation_text="Umbral Levas 0.0001")
                fig.add_vline(x=CHATTER_APOYOS_THRESHOLD, line_dash="dot", annotation_text="Umbral Apoyos 0.00008")
                st.plotly_chart(fig, use_container_width=True)

    with tab5:
        selected = st.selectbox("Selecciona una característica para profundizar", sorted(view["Caracteristica"].dropna().unique()))
        d = view[view["Caracteristica"] == selected].copy()
        if not d.empty:
            col1, col2 = st.columns(2)
            with col1:
                d2 = d.groupby("Resultado").size().reset_index(name="Cantidad")
                fig = px.pie(d2, names="Resultado", values="Cantidad", hole=.55, title=f"{selected}: OK / NOK")
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                fig = px.box(d, x="Resultado", y="Medicion", points="outliers", title=f"Distribución de medición — {selected}")
                if d["Spec Min"].notna().any():
                    fig.add_hline(y=float(d["Spec Min"].dropna().iloc[0]), line_dash="dash", annotation_text="Spec Min")
                if d["Spec Max"].notna().any():
                    fig.add_hline(y=float(d["Spec Max"].dropna().iloc[0]), line_dash="dash", annotation_text="Spec Max")
                st.plotly_chart(fig, use_container_width=True)

            location_col = "Leva" if d["Leva"].notna().any() else "Apoyo"
            loc = d[d.Resultado=="Nok"].groupby(location_col).size().reset_index(name="NOK").sort_values("NOK", ascending=False)
            if not loc.empty:
                fig = px.bar(loc, x=location_col, y="NOK", title=f"{selected}: NOK por {location_col}")
                st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 🔍 Registros NOK")
    if not nok.empty:
        st.dataframe(nok.sort_values(["Caracteristica", "Pieza"]), use_container_width=True, height=350)


# ============================================================
# UI PRINCIPAL
# ============================================================
with st.sidebar:
    st.markdown("## ⚙️ Configuración")
    st.metric("Máximo de PDFs por ejecución", f"{MAX_INPUT_PDFS:,}")
    workers = st.slider("Workers paralelos", min_value=1, max_value=MAX_WORKERS, value=4)
    st.caption("Puedes seleccionar hasta 1,500 PDFs de una vez. Cada PDF se procesa y guarda individualmente para conservar todas sus mediciones.")
    st.divider()
    st.markdown("### Reglas activas")
    st.code("Levas (5) 301-400 UPR ≥ 0.0001 → NOK\nChatter Apoyos → 0.00008", language="text")
    if st.button("🗑️ Nueva sesión / limpiar resultados", use_container_width=True):
        reset_workspace()

st.markdown("### 1️⃣ Cargar archivos")
st.info("Modo individual: puedes seleccionar hasta 1,500 PDFs en una sola ejecución. Cada PDF conserva sus registros, mediciones, Pieza, Leva/Apoyo y Resultado; después todos se consolidan en el reporte final. También acepta uno o varios ZIP y mezcla PDF + ZIP.")

uploaded_files = st.file_uploader(
    "Selecciona PDFs y/o ZIPs",
    type=["pdf", "zip"],
    accept_multiple_files=True,
    help="Puedes seleccionar hasta 1,500 PDFs. Cada PDF se procesa individualmente y sus resultados se conservan antes de consolidar el reporte. Los ZIP se extraen y procesan en disco temporal.",
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
col1, col2 = st.columns(2)
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
            file_name="Reporte_NoK.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

if error_count:
    with st.expander(f"⚠️ Ver errores ({error_count})"):
        rows = conn.execute("SELECT filename,piece,error,processed_at FROM files WHERE status='ERROR' ORDER BY piece").fetchall()
        st.dataframe(pd.DataFrame(rows, columns=["Archivo","Pieza","Error","Fecha"]), use_container_width=True)

st.caption(f"Reporte NoK {APP_VERSION} · procesamiento individual por PDF · SHA-256 · {datetime.now().strftime('%Y-%m-%d %H:%M')}")

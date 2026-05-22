import os
import re
import jwt
import hashlib
import anthropic
import chromadb
import pandas as pd
import sqlite3
import sqlparse
import pyodbc
import xml.etree.ElementTree as ET
import bcrypt
import httpx
import base64
import io
import time
from fastapi import FastAPI, HTTPException, Depends, Request
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from datetime import datetime, timedelta
from pathlib import Path
from filelock import FileLock
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

app = FastAPI()

# ── Rate Limiting ─────────────────────────────────────────────────────────────
limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

# ── CORS restringido ──────────────────────────────────────────────────────────
ALLOWED_ORIGINS = os.getenv("ALLOWED_ORIGINS", "http://localhost,https://hurpayai.ximery.com,http://143.198.125.134:8000,http://hurpay.com.pe,https://hurpay.com.pe,http://www.hurpay.com.pe,https://www.hurpay.com.pe").split(",")
app.add_middleware(CORSMiddleware, allow_origins=ALLOWED_ORIGINS, allow_credentials=True,
                   allow_methods=["GET", "POST", "PUT", "DELETE"], allow_headers=["*"])

JWT_SECRET        = os.getenv("JWT_SECRET", "hurpay-secreto-2024")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "").strip()
NOMBRE_EMPRESA    = os.getenv("NOMBRE_EMPRESA", "Hurpay")
USUARIOS_FILE     = "/app/usuarios.xml"
USUARIOS_LOCK     = "/app/usuarios.xml.lock"

SQL_SERVER   = os.getenv("SQL_SERVER",   "p1426.use1.mysecurecloudhost.com")
SQL_DATABASE = os.getenv("SQL_DATABASE", "javierd2_CapitalSoft")
SQL_USER     = os.getenv("SQL_USER",     "javierd2_capitalsoftia")
SQL_PASSWORD = os.getenv("SQL_PASSWORD", "sasasa0")

EXCEL_MODELO  = "/app/modelos/Nómina_IA_2_-_prueba.xlsx"
AZURE_API_URL = "https://hurpayai.ximery.com/api"

TABLAS_CON_DOCUMENTO = {
    "CONContratoTrabajador", "CONProrrogaContrato",
    "TRABoletaPago", "TRAComprobantePostLBS",
    "TRACertificadoCTS", "TRACertificadoLibreDisponibilidad",
    "TRACertificadoQuinta", "TRACertificadoTrabajo",
    "TRACertificadoUtilidades", "TRAComprobanteLS",
}

# ── Caché de consultas en memoria (TTL: 1 hora) ───────────────────────────────
_CACHE_CONSULTAS: dict[str, tuple[float, dict]] = {}
_CACHE_TTL_SEG = 3600  # 1 hora

def _cache_key(usuario: str, mensaje: str) -> str:
    return hashlib.md5(f"{usuario}::{mensaje.strip().lower()}".encode()).hexdigest()

def _cache_get(key: str) -> dict | None:
    entry = _CACHE_CONSULTAS.get(key)
    if entry and (time.time() - entry[0]) < _CACHE_TTL_SEG:
        return entry[1]
    if entry:
        del _CACHE_CONSULTAS[key]
    return None

def _cache_set(key: str, valor: dict):
    if len(_CACHE_CONSULTAS) >= 500:
        mas_viejo = min(_CACHE_CONSULTAS, key=lambda k: _CACHE_CONSULTAS[k][0])
        del _CACHE_CONSULTAS[mas_viejo]
    _CACHE_CONSULTAS[key] = (time.time(), valor)


# ── Azure PDF ──────────────────────────────────────────────────────────────────
async def obtener_token_azure(username: str, password: str, api_url: str = None) -> str:
    base_url = (api_url or AZURE_API_URL).rstrip("/")
    async with httpx.AsyncClient(timeout=15, verify=False) as client:
        r = await client.post(f"{base_url}/auth/login",
                              headers={"env": "CAPLAN"},
                              json={"Username": username, "Password": password})
        r.raise_for_status()
        return r.json()["ReturnValue"]["Token"]

async def descargar_pdf_azure(tabla: str, internal_id: int, token: str, api_url: str = None) -> bytes:
    base_url = (api_url or AZURE_API_URL).rstrip("/")
    url = f"{base_url}/forms/{tabla}/File/Documento/{internal_id}"
    async with httpx.AsyncClient(timeout=15, verify=False) as client:
        r = await client.get(url, headers={"env": "CAPLAN", "token": token})
        r.raise_for_status()
        return base64.b64decode(r.json()["ReturnValue"])


# ── Hash MD5 del Excel ────────────────────────────────────────────────────────
def _hash_excel(path: str) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


# ── Carga completa del modelo ─────────────────────────────────────────────────
def cargar_todo() -> dict:
    """
    Lee las 7 pestañas del Excel Nómina_IA_2_-_prueba:
      1. Condiciones Generales → reglas del sistema (columna A, sin header)
      2. Equivalencia Tablas   → tabla PowerBI <-> tabla SQL, descripcion, CTipId
      3. Equivalencia Columnas → campo PowerBI <-> columna SQL
      4. Relaciones            → JOINs documentados con columna "JOIN SQL" exacta
      5. Cadenas de Join       → cadenas de joins paso a paso
      6. Tablas Lista          → valores de CAPLAN_DetTipo por CTipId
      7. Sinónimos             → términos laborales peruanos
    """
    xls = pd.ExcelFile(EXCEL_MODELO)
    sheet_names = xls.sheet_names

    # Validar que la pestaña Condiciones Generales existe
    if "Condiciones Generales" not in sheet_names:
        raise RuntimeError(
            f"❌ ERROR: La pestaña 'Condiciones Generales' no se encontró en el Excel. "
            f"Pestañas disponibles: {sheet_names}"
        )

    def _leer(nombre, **kw):
        return pd.read_excel(xls, sheet_name=nombre, **kw) if nombre in sheet_names else pd.DataFrame()

    # ── Condiciones Generales: sin header, columna A (índice 0) ──────────────
    cond_gen   = _leer("Condiciones Generales", header=None)

    # ── Resto de pestañas con header en fila 0 ────────────────────────────────
    eq_tablas  = _leer("Equivalencia Tablas")
    eq_cols    = _leer("Equivalencia Columnas")
    relaciones = _leer("Relaciones")
    joins      = _leer("Cadenas de Join")
    t_lista    = _leer("Tablas Lista")
    sinonimos  = _leer("Sinónimos")

    # ── Limpiar Equivalencia Tablas ───────────────────────────────────────────
    # Eliminar columnas extra vacías (Unnamed: 5, 6, 7, 8...)
    if not eq_tablas.empty:
        cols_validas = [c for c in eq_tablas.columns if not str(c).startswith("Unnamed")]
        eq_tablas = eq_tablas[cols_validas]
        eq_tablas = eq_tablas.dropna(subset=["Tabla SQL"])
        eq_tablas = eq_tablas[~eq_tablas["Tabla SQL"].astype(str).str.strip().isin(["", "nan"])]
        eq_tablas = eq_tablas.where(pd.notna(eq_tablas), other="")

    # ── Limpiar Equivalencia Columnas ─────────────────────────────────────────
    if not eq_cols.empty:
        eq_cols = eq_cols.dropna(subset=["Columna SQL"])
        eq_cols = eq_cols[~eq_cols["Columna SQL"].astype(str).str.strip().isin(["", "nan"])]
        eq_cols = eq_cols.where(pd.notna(eq_cols), other="")

    # ── Limpiar Relaciones ────────────────────────────────────────────────────
    if not relaciones.empty:
        relaciones = relaciones.dropna(subset=["JOIN SQL"])
        relaciones = relaciones[~relaciones["JOIN SQL"].astype(str).str.strip().isin(["", "nan"])]
        relaciones = relaciones.where(pd.notna(relaciones), other="")

    for df in (joins, t_lista, sinonimos):
        if not df.empty:
            df.dropna(how="all", inplace=True)

    # ── Mapa tabla BI → CTipId (para filtros CAPLAN_DetTipo) ─────────────────
    bi_to_ctipid: dict[str, str] = {}
    for _, row in eq_tablas.iterrows():
        bi   = str(row.get("Tabla PowerBI", "")).strip()
        ctip = str(row.get("CTipId", "")).strip()
        if bi and ctip and ctip not in ("", "nan"):
            bi_to_ctipid[bi] = ctip

    # ── Resolver relaciones desde pestaña Relaciones ──────────────────────────
    relaciones_resueltas: list[dict] = []

    if not relaciones.empty:
        for _, r in relaciones.iterrows():
            from_bi  = str(r.get("Tabla FROM (BI)",  "")).strip()
            from_sql = str(r.get("Tabla FROM (SQL)", "")).strip()
            to_bi    = str(r.get("Tabla TO (BI)",    "")).strip()
            to_sql   = str(r.get("Tabla TO (SQL)",   "")).strip()
            join_sql = str(r.get("JOIN SQL",          "")).strip()

            col_from_bi  = str(r.get("Columna FROM (BI)",  "")).strip()
            col_from_sql = str(r.get("Columna FROM (SQL)", "")).strip()
            col_to_bi    = str(r.get("Columna TO (BI)",    "")).strip()
            col_to_sql   = str(r.get("Columna TO (SQL)",   "")).strip()

            if not from_sql or not to_sql or not join_sql:
                continue
            if from_sql in ("nan", "") or to_sql in ("nan", ""):
                continue

            ctipid_filtro = None
            if to_sql == "CAPLAN_DetTipo" and to_bi:
                ct = bi_to_ctipid.get(to_bi, "")
                if ct and ct not in ("", "nan"):
                    try:
                        ctipid_filtro = int(ct)
                    except ValueError:
                        pass

            relaciones_resueltas.append({
                "from_sql":      from_sql,
                "col_from_sql":  col_from_sql,
                "to_sql":        to_sql,
                "col_to_sql":    col_to_sql,
                "from_bi":       from_bi,
                "col_from_bi":   col_from_bi,
                "to_bi":         to_bi,
                "col_to_bi":     col_to_bi,
                "join_sql":      join_sql,
                "ctipid_filtro": ctipid_filtro,
            })

    # ── Reglas del sistema: columna A (índice 0), todas las filas no nulas ────
    reglas = ""
    if not cond_gen.empty:
        reglas = "\n".join(
            cond_gen.iloc[:, 0]          # columna A
            .dropna()
            .astype(str)
            .str.strip()
            .tolist()
        )

    # Validar que las reglas no quedaron vacías
    assert reglas, (
        "❌ ERROR: La pestaña 'Condiciones Generales' está vacía o no tiene datos en la columna A. "
        "Verifica el contenido del Excel."
    )

    # ── Texto compacto de las pestañas para el system prompt ─────────────────
    def _df_pipe(df, cols):
        """Convierte un DataFrame a texto pipe-separated con solo las columnas indicadas."""
        cols_ok = [c for c in cols if c in df.columns]
        sub = df[cols_ok].fillna("").astype(str)
        header = "|".join(cols_ok)
        rows = "\n".join("|".join(row[c].strip() for c in cols_ok) for _, row in sub.iterrows())
        return header + "\n" + rows

    # Equivalencia Tablas compacta
    et_prompt = _df_pipe(eq_tablas, ["Tabla PowerBI", "Tabla SQL", "Categoría", "CTipId", "Documentación"])

    # Equivalencia Columnas compacta
    ec_prompt = _df_pipe(eq_cols, ["Tabla SQL", "Campo PowerBI", "Columna SQL"])

    # Relaciones compacta
    rel_cols = ["Tabla FROM (SQL)", "Columna FROM (SQL)", "Tabla TO (SQL)", "Columna TO (SQL)", "JOIN SQL"]
    rel_prompt = _df_pipe(relaciones, rel_cols)

    # Tablas Lista compacta
    tl_cols = [c for c in ["Tabla PowerBI", "Tabla SQL", "CtipId", "DtipNombre", "DtipCodigo", "DtipId", "InternalId"] if c in t_lista.columns]
    tl_clean = t_lista.dropna(subset=["DtipNombre"] if "DtipNombre" in t_lista.columns else ([tl_cols[0]] if tl_cols else []))
    tl_prompt = _df_pipe(tl_clean, tl_cols)

    # Sinónimos compacta
    sin_cols = [c for c in ["Término estándar", "Sinónimo en Perú", "Notas de uso"] if c in sinonimos.columns]
    sin_prompt = _df_pipe(sinonimos.dropna(how="all"), sin_cols)

    prompt_modelo = f"""=== CONDICIONES GENERALES (reglas del sistema — aplicar siempre en cada consulta) ===
{reglas}

=== EQUIVALENCIA TABLAS (fuente de verdad: qué tabla SQL usar) ===
{et_prompt}

=== EQUIVALENCIA COLUMNAS (fuente de verdad: únicas columnas válidas para SELECT) ===
{ec_prompt}

=== RELACIONES (fuente de verdad: JOINs exactos entre tablas) ===
{rel_prompt}

=== TABLAS LISTA — valores de CAPLAN_DetTipo (fuente de verdad: valores para filtros WHERE) ===
{tl_prompt}

=== SINÓNIMOS LABORALES PERÚ ===
{sin_prompt}"""

    print(
        f"✅ Excel cargado correctamente:\n"
        f"   • Condiciones Generales: {len([l for l in reglas.splitlines() if l.strip()])} reglas\n"
        f"   • Equivalencia Tablas:   {len(eq_tablas)} tablas\n"
        f"   • Equivalencia Columnas: {len(eq_cols)} columnas\n"
        f"   • Relaciones:            {len(relaciones_resueltas)} relaciones resueltas\n"
        f"   • Tablas Lista:          {len(t_lista)} valores\n"
        f"   • Sinónimos:             {len(sinonimos)} términos"
    )

    return {
        "eq_tablas":            eq_tablas,
        "eq_cols":              eq_cols,
        "relaciones":           relaciones,
        "joins":                joins,
        "t_lista":              t_lista,
        "reglas":               reglas,
        "sinonimos":            sinonimos,
        "relaciones_resueltas": relaciones_resueltas,
        "prompt_modelo":        prompt_modelo,
        "hash_excel":           _hash_excel(EXCEL_MODELO),
    }


# ── Indexar en ChromaDB ───────────────────────────────────────────────────────
def indexar_modelo_en_chroma():
    """
    Un documento por tabla SQL única.
    Re-indexa solo si cambió el MD5 del Excel.
    """
    col = chroma_client.get_or_create_collection("modelo_datos")
    hash_actual = MODELO["hash_excel"]

    hash_guardado = None
    try:
        ctrl = col.get(ids=["__control__"])
        if ctrl and ctrl["metadatas"]:
            hash_guardado = ctrl["metadatas"][0].get("hash_excel")
    except Exception:
        pass

    n_esperado = len(MODELO["eq_tablas"]) + 2
    if hash_guardado == hash_actual and col.count() >= n_esperado:
        return col

    try:
        chroma_client.delete_collection("modelo_datos")
    except Exception:
        pass
    col = chroma_client.get_or_create_collection("modelo_datos")

    eq_tablas            = MODELO["eq_tablas"]
    eq_cols              = MODELO["eq_cols"]
    joins                = MODELO["joins"]
    t_lista              = MODELO["t_lista"]
    relaciones_resueltas = MODELO["relaciones_resueltas"]

    def _filtrar_joins(df: pd.DataFrame, tabla_sql: str) -> pd.DataFrame:
        if df.empty:
            return df
        corto = tabla_sql.replace("CAPLAN_", "")
        mask  = pd.Series(False, index=df.index)
        for c in df.select_dtypes(include="object").columns:
            s = df[c].astype(str)
            mask |= s.str.contains(re.escape(tabla_sql), case=False, na=False)
            mask |= s.str.contains(
                r"(?i)(?:^|\b)" + re.escape(corto) + r"(?:\b|$)", regex=True, na=False)
        return df[mask]

    def _filtrar_lista_valores(df: pd.DataFrame, tabla_sql: str, ctipid: str) -> pd.DataFrame:
        if df.empty:
            return df
        if tabla_sql == "CAPLAN_DetTipo" and ctipid and ctipid not in ("", "nan"):
            try:
                ct_int = int(ctipid)
                if "CtipId" in df.columns:
                    return df[df["CtipId"].astype(str) == str(ct_int)]
            except (ValueError, TypeError):
                pass
        mask = pd.Series(False, index=df.index)
        for c in df.select_dtypes(include="object").columns:
            s = df[c].astype(str)
            mask |= s.str.contains(re.escape(tabla_sql), case=False, na=False)
        return df[mask]

    def _df_txt(df):
        return df.to_string(index=False) if not df.empty else ""

    docs, metas, ids = [], [], []
    ids_vistos: set[str] = set()

    for _, row in eq_tablas.iterrows():
        tabla_sql = str(row.get("Tabla SQL", "")).strip()
        tabla_bi  = str(row.get("Tabla PowerBI", "")).strip()
        categoria = str(row.get("Categoría", "")).strip()
        ctipid    = str(row.get("CTipId", "")).strip()
        doc_texto = str(row.get("Documentación", "")).strip()

        if not tabla_sql or tabla_sql in ("nan", ""):
            continue

        doc_id = f"tabla_{tabla_sql}_{ctipid}" if (tabla_sql == "CAPLAN_DetTipo" and ctipid not in ("", "nan")) else f"tabla_{tabla_sql}"

        if doc_id in ids_vistos:
            continue
        ids_vistos.add(doc_id)

        cols_tabla = eq_cols[eq_cols["Tabla SQL"].astype(str).str.strip() == tabla_sql]

        cols_lineas = []
        for _, c in cols_tabla.iterrows():
            bi  = str(c.get("Campo PowerBI", "")).strip()
            sql = str(c.get("Columna SQL", "")).strip()
            if not sql or sql == "nan":
                continue
            cols_lineas.append(
                f"  {bi} → {sql}" if (bi and bi not in ("nan", "") and bi != sql) else f"  {sql}"
            )
        cols_texto = "\n".join(cols_lineas) if cols_lineas else "Sin columnas documentadas"

        rels_lineas = []
        for rel in relaciones_resueltas:
            if tabla_sql == "CAPLAN_DetTipo" and ctipid not in ("", "nan"):
                if rel["to_sql"] == tabla_sql:
                    filtro_ctip = str(rel.get("ctipid_filtro", "") or "")
                    if filtro_ctip and filtro_ctip != ctipid:
                        continue

            filtro = f" AND {rel['to_sql']}.CTipId={rel['ctipid_filtro']}" if rel.get("ctipid_filtro") else ""
            if rel["from_sql"] == tabla_sql:
                rels_lineas.append(
                    f"  → {rel['to_bi']} [{rel['to_sql']}]\n"
                    f"    JOIN: {rel['join_sql']}{filtro}"
                )
            elif rel["to_sql"] == tabla_sql:
                rels_lineas.append(
                    f"  ← {rel['from_bi']} [{rel['from_sql']}]\n"
                    f"    JOIN: {rel['join_sql']}"
                )
        rels_texto = "\n".join(rels_lineas) if rels_lineas else "Sin relaciones registradas"

        joins_rel   = _filtrar_joins(joins, tabla_sql)
        joins_texto = _df_txt(joins_rel) if not joins_rel.empty else "Sin joins documentados"

        lista_rel  = _filtrar_lista_valores(t_lista, tabla_sql, ctipid)
        lista_txt  = _df_txt(lista_rel) if not lista_rel.empty else "No aplica"

        tiene_pdf = (tabla_sql.replace("CAPLAN_", "") in TABLAS_CON_DOCUMENTO
                     or tabla_sql in TABLAS_CON_DOCUMENTO)

        doc = f"""TABLA SQL: {tabla_sql}
NOMBRE EN POWERBI: {tabla_bi}
CATEGORÍA: {categoria}
CTIPID: {ctipid}
DESCRIPCIÓN: {doc_texto}
TIENE PDF: {"Sí — incluir InternalId en SELECT" if tiene_pdf else "No"}

COLUMNAS VÁLIDAS PARA SELECT (Campo BI → Columna SQL):
{cols_texto}

RELACIONES (usar JOIN SQL exacto — no inventar columnas):
{rels_texto}

CADENAS DE JOIN DOCUMENTADAS:
{joins_texto}

VALORES DE LISTA (para filtros WHERE):
{lista_txt}
"""
        docs.append(doc)
        metas.append({"tipo": "tabla", "tabla_sql": tabla_sql,
                      "tabla_bi": tabla_bi, "categoria": categoria})
        ids.append(doc_id)

    # ── Documento global: reglas + sinónimos + mapa completo de relaciones ─────
    mapa_rels = "\n".join(
        "  {join_sql}{filtro}  ({from_bi} → {to_bi})".format(
            join_sql=r["join_sql"],
            filtro=f" [CTipId={r['ctipid_filtro']}]" if r.get("ctipid_filtro") else "",
            from_bi=r.get("from_bi", ""),
            to_bi=r.get("to_bi", ""),
        )
        for r in relaciones_resueltas
        if r.get("from_sql") and r.get("to_sql")
    )

    docs.append(f"""REGLAS GENERALES DEL SISTEMA HURPAY:
{MODELO['reglas']}

SINÓNIMOS Y TÉRMINOS LABORALES PERUANOS:
{_df_txt(MODELO['sinonimos'])}

MAPA COMPLETO DE RELACIONES SQL ({len(relaciones_resueltas)} relaciones con JOIN exacto):
{mapa_rels}

CADENAS DE JOIN COMPLETAS (referencia de navegación entre tablas):
{_df_txt(MODELO['joins'])}
""")
    metas.append({"tipo": "reglas", "tabla_sql": "GLOBAL",
                  "tabla_bi": "GLOBAL", "categoria": "GLOBAL"})
    ids.append("reglas_generales")

    docs.append(f"Control de versión. Hash Excel: {hash_actual}")
    metas.append({"tipo": "control", "tabla_sql": "CONTROL",
                  "tabla_bi": "CONTROL", "categoria": "CONTROL",
                  "hash_excel": hash_actual})
    ids.append("__control__")

    batch = 50
    for i in range(0, len(docs), batch):
        col.add(documents=docs[i:i+batch], metadatas=metas[i:i+batch], ids=ids[i:i+batch])

    return col


# ── Contexto relevante para la pregunta ──────────────────────────────────────

# ── SQLITE - USUARIOS ───────────────────────────────────────────────────────
def init_log_consultas():
    """Crea la tabla log_consultas si no existe."""
    conn = sqlite3.connect('/app/usuarios.db')
    conn.execute("""
        CREATE TABLE IF NOT EXISTS log_consultas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha DATE DEFAULT (DATE('now')),
            hora TIME DEFAULT (TIME('now')),
            usuario TEXT,
            modelo TEXT,
            input_tokens INTEGER DEFAULT 0,
            cache_read_tokens INTEGER DEFAULT 0,
            cache_write_tokens_5m INTEGER DEFAULT 0,
            cache_write_tokens_1h INTEGER DEFAULT 0,
            output_tokens INTEGER DEFAULT 0,
            costo_usd REAL DEFAULT 0,
            pregunta TEXT
        )
    """)
    conn.commit()
    conn.close()

def guardar_log_consulta(usuario: str, modelo: str, usage, pregunta: str):
    """Guarda el log de una consulta con su consumo de tokens y costo."""
    try:
        # Calcular costo según modelo
        sonnet_input = 3.00
        sonnet_cache_read = 0.30
        sonnet_cache_write = 3.75
        sonnet_output = 15.00
        haiku_input = 0.80
        haiku_cache_read = 0.08
        haiku_cache_write = 1.00
        haiku_output = 4.00

        if 'sonnet' in modelo.lower():
            p_in = sonnet_input
            p_cr = sonnet_cache_read
            p_cw = sonnet_cache_write
            p_out = sonnet_output
        else:
            p_in = haiku_input
            p_cr = haiku_cache_read
            p_cw = haiku_cache_write
            p_out = haiku_output

        input_tokens = getattr(usage, 'input_tokens', 0)
        output_tokens = getattr(usage, 'output_tokens', 0)
        cache_read = getattr(usage, 'cache_read_input_tokens', 0)
        cache_write_tokens = getattr(usage, 'cache_creation_input_tokens', 0)

        # Detectar si es 5m o 1h (1h si es el primer write grande)
        cache_write_5m = cache_write_tokens if cache_write_tokens < 1000 else 0
        cache_write_1h = cache_write_tokens if cache_write_tokens >= 1000 else 0

        costo = (
            (input_tokens * p_in / 1_000_000) +
            (cache_read * p_cr / 1_000_000) +
            (cache_write_tokens * p_cw / 1_000_000) +
            (output_tokens * p_out / 1_000_000)
        )

        conn = sqlite3.connect('/app/usuarios.db')
        conn.execute("""
            INSERT INTO log_consultas 
            (usuario, modelo, input_tokens, cache_read_tokens, cache_write_tokens_5m, cache_write_tokens_1h, output_tokens, costo_usd, pregunta)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (usuario, modelo, input_tokens, cache_read, cache_write_5m, cache_write_1h, output_tokens, round(costo, 6), pregunta[:500]))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Error guardando log: {e}")

# Inicializar tabla al arrancar
init_log_consultas()

def init_prompts():
    """Crea las tablas de prompts favoritos y categorías si no existen."""
    conn = sqlite3.connect("/app/usuarios.db")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS prompt_categorias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL,
            nombre_categoria TEXT NOT NULL,
            fecha TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS prompts_favoritos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL,
            categoria_id INTEGER NOT NULL,
            nombre TEXT NOT NULL,
            prompt TEXT NOT NULL,
            fecha TEXT
        )
    """)
    conn.commit()
    conn.close()

init_prompts()

def obtener_usuario_sqlite(usuario_codigo: str) -> dict:
    """
    Obtiene un usuario de la BD SQLite.
    Retorna dict con credenciales SQL dinámicas por usuario.
    """
    try:
        conn = sqlite3.connect('/app/usuarios.db')
        cursor = conn.cursor()
        
        cursor.execute(
            "SELECT usuario, nombre_usuario, password, servidor_bd, base_datos, api_url, rol, situacion, usuario_bd, password_bd FROM usuarios WHERE usuario=?",
            (usuario_codigo,)
        )
        
        resultado = cursor.fetchone()
        conn.close()
        
        if resultado:
            return {
                'usuario': resultado[0],
                'nombre_usuario': resultado[1],
                'password': resultado[2],
                'servidor_bd': resultado[3],
                'base_datos': resultado[4],
                'api_url': resultado[5],
                'rol': resultado[6],
                'situacion': resultado[7],
                'usuario_bd': resultado[8] or '',
                'password_bd': resultado[9] or ''
            }
        return None
    except Exception as e:
        print(f"Error leyendo SQLite: {e}")
        return None


def modelo_relevante(pregunta: str, n_tablas: int = 8) -> str:
    try:
        n_safe = min(n_tablas, max(coleccion_modelo.count() - 2, 1))
        res = coleccion_modelo.query(
            query_texts=[pregunta], n_results=n_safe, where={"tipo": "tabla"})
        tablas_relevantes = "\n\n---\n\n".join(res["documents"][0])
    except Exception:
        tablas_relevantes = "(No se encontraron tablas relevantes)"

    try:
        rg = coleccion_modelo.get(ids=["reglas_generales"])
        reglas_texto = rg["documents"][0] if rg["documents"] else MODELO["reglas"]
    except Exception:
        reglas_texto = MODELO["reglas"]

    kw_reglas = {"regla", "condición", "condicion", "sinónimo", "sinonimo",
                 "restricción", "restriccion", "relacion", "relación", "join"}
    if not any(kw in pregunta.lower() for kw in kw_reglas):
        lineas = reglas_texto.strip().split("\n")
        reglas_texto = "\n".join(lineas[:30])
        if len(lineas) > 30:
            reglas_texto += f"\n... ({len(lineas) - 30} líneas omitidas)"

    return f"""{reglas_texto}

=== TABLAS Y COLUMNAS RELEVANTES ===
{tablas_relevantes}
"""


# ── SQL Server ────────────────────────────────────────────────────────────────
def get_conexion_sql(base_datos: str = None, servidor: str = None, usuario_bd: str = None, password_bd: str = None):
    db = base_datos or SQL_DATABASE
    srv = servidor or SQL_SERVER
    uid = usuario_bd or SQL_USER
    pwd = password_bd or SQL_PASSWORD
    conn_str = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={srv};DATABASE={db};"
        f"UID={uid};PWD={pwd};"
        f"TrustServerCertificate=yes;Encrypt=yes;"
    )
    return pyodbc.connect(conn_str, timeout=15)

def ejecutar_sql(query: str, base_datos: str = None, servidor: str = None, usuario_bd: str = None, password_bd: str = None):
    try:
        conn = get_conexion_sql(base_datos, servidor, usuario_bd, password_bd)
        cursor = conn.cursor()
        cursor.execute(query)
        columnas = [desc[0] for desc in cursor.description]
        filas = cursor.fetchmany(500)
        conn.close()
        return {"columnas": columnas, "filas": [list(f) for f in filas], "total": len(filas)}
    except Exception as e:
        return {"error": str(e)}


def validar_sql_seguro_mejorado(query: str) -> tuple[bool, str]:
    """
    Valida SQL con SQLParse para máxima seguridad.
    Retorna: (es_seguro, mensaje)
    """
    try:
        parsed = sqlparse.parse(query)
        
        # Comandos peligrosos ABSOLUTAMENTE PROHIBIDOS
        peligrosos = ['DROP', 'DELETE', 'INSERT', 'UPDATE', 'CREATE', 'ALTER', 
                     'TRUNCATE', 'EXEC', 'EXECUTE', 'MERGE', 'GRANT', 'REVOKE']
        
        for statement in parsed:
            # Obtener el primer token (comando SQL)
            first_token = statement.token_first(skip_ws=True, skip_cm=True)
            
            if first_token and first_token.ttype is sqlparse.tokens.Keyword:
                if first_token.value.upper() in peligrosos:
                    return False, f"⚠️ Comando '{first_token.value.upper()}' no permitido"
        
        return True, "OK"
    except Exception as e:
        return False, f"Error al validar SQL: {str(e)}"


def detectar_sql_peligroso(query: str) -> bool:
    q = query.upper()
    return any(p in q for p in ["INSERT","UPDATE","DELETE","DROP","TRUNCATE",
                                  "ALTER","CREATE","EXEC","EXECUTE","MERGE"])


# ── Usuarios XML ──────────────────────────────────────────────────────────────
def _xml_a_lista(root) -> list:
    return [{
        "usuario":    (u.findtext("usuario")    or "").strip(),
        "nombre_usuario": (u.findtext("nombre_usuario") or "").strip(),
        "password":  (u.findtext("password")  or "").strip(),
        "base_datos":     (u.findtext("base_datos")     or "").strip(),
        "api_url":       (u.findtext("api_url")       or "").strip(),
        "situacion":     (u.findtext("situacion")     or "inactivo").strip(),
        "rol":           (u.findtext("rol")           or "usuario").strip(),
    } for u in root.findall("usuario")]

def _lista_a_xml(usuarios: list) -> ET.Element:
    root = ET.Element("usuarios")
    for u in usuarios:
        elem = ET.SubElement(root, "usuario")
        for campo in ["usuario","nombre_usuario","password",
                      "base_datos","api_url","situacion","rol"]:
            ET.SubElement(elem, campo).text = u.get(campo, "")
    return root

def cargar_usuarios() -> list:
    path = Path(USUARIOS_FILE)
    if not path.exists():
        pwd_hash = bcrypt.hashpw("admin123".encode(), bcrypt.gensalt()).decode()
        usuarios = [{"usuario": "admin", "nombre_usuario": "Administrador",
                     "password": pwd_hash, "base_datos": SQL_DATABASE,
                     "api_url": "/pdfs/", "situacion": "activo", "rol": "administrador"}]
        guardar_usuarios(usuarios)
        return usuarios
    with FileLock(USUARIOS_LOCK):
        return _xml_a_lista(ET.parse(USUARIOS_FILE).getroot())

def guardar_usuarios(usuarios: list):
    root = _lista_a_xml(usuarios)
    ET.indent(root, space="  ")
    with FileLock(USUARIOS_LOCK):
        tree = ET.ElementTree(root)
        ET.indent(tree.getroot(), space="  ")
        tree.write(USUARIOS_FILE, encoding="utf-8", xml_declaration=True)



# ── ChromaDB + Anthropic ──────────────────────────────────────────────────────
chroma_client = chromadb.EphemeralClient()
coleccion        = chroma_client.get_or_create_collection("manuales")
cliente_ai       = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

MODELO           = cargar_todo()
coleccion_modelo = indexar_modelo_en_chroma()

# ── Indexar PDFs de manuales en ChromaDB ─────────────────────────────────────
def indexar_manuales():
    """
    Indexa los PDFs de /app/manuales/ en la coleccion 'manuales' de ChromaDB.
    Una pagina = un documento. Re-indexa solo si cambia el set de PDFs.
    """
    import PyPDF2
    from pathlib import Path

    CARPETA = "/app/manuales"
    pdfs = sorted(Path(CARPETA).glob("*.pdf"))

    if not pdfs:
        print("No se encontraron PDFs en /app/manuales/")
        return

    nombres_pdfs = sorted([p.name for p in pdfs])
    control_id = "control_manuales"
    try:
        ctrl = coleccion.get(ids=[control_id])
        if ctrl and ctrl["metadatas"] and ctrl["metadatas"][0].get("pdfs") == str(nombres_pdfs):
            print(f"Manuales ya indexados ({coleccion.count()} fragmentos)")
            return
    except Exception:
        pass

    try:
        existing = coleccion.get()
        if existing and existing["ids"]:
            coleccion.delete(ids=existing["ids"])
    except Exception:
        pass

    total_paginas = 0
    for ruta_pdf in pdfs:
        nombre = ruta_pdf.name
        print(f"  Indexando: {nombre}")
        try:
            with open(ruta_pdf, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                for i, pagina in enumerate(reader.pages):
                    texto = pagina.extract_text()
                    if texto and len(texto.strip()) > 50:
                        doc_id = f"{nombre}_p{i+1}"
                        coleccion.add(
                            documents=[texto],
                            metadatas=[{"fuente": nombre, "pagina": i + 1}],
                            ids=[doc_id]
                        )
                        total_paginas += 1
        except Exception as e:
            print(f"  Error indexando {nombre}: {e}")

    try:
        coleccion.add(
            documents=["control"],
            metadatas=[{"pdfs": str(nombres_pdfs)}],
            ids=[control_id]
        )
    except Exception:
        pass

    print(f"Manuales indexados: {len(pdfs)} PDFs, {total_paginas} fragmentos")


indexar_manuales()


# ── Pydantic ──────────────────────────────────────────────────────────────────
class LoginRequest(BaseModel):
    usuario: str
    password: str

class ChatRequest(BaseModel):
    mensaje: str
    historial: list = []
    
    class Config:
        extra = "allow"

class UsuarioRequest(BaseModel):
    usuario:        str
    nombre_usuario: str
    password:       str = ""
    servidor_bd:    str = ""
    base_datos:     str = ""
    api_url:        str = ""
    usuario_bd:     str = ""
    password_bd:    str = ""
    situacion:     str = "activo"
    rol:           str = "usuario"

class CategoriaRequest(BaseModel):
    nombre_categoria: str

class PromptFavoritoRequest(BaseModel):
    categoria_id: int
    nombre: str
    prompt: str


# ── JWT ───────────────────────────────────────────────────────────────────────
def crear_token_dinamico(usuario: dict, password_plano: str = ""):
    """
    Crea JWT con credenciales SQL dinámicas por usuario.
    """
    payload = {
        "sub": usuario["usuario"],
        "rol": usuario["rol"],
        "servidor_bd": usuario["servidor_bd"],
        "base_datos": usuario["base_datos"],
        "api_url": usuario["api_url"],
        "pwd_azure": password_plano,
        "usuario_bd": usuario.get("usuario_bd", ""),
        "password_bd": usuario.get("password_bd", ""),
        "exp": datetime.utcnow() + timedelta(hours=8),
    }
    return jwt.encode(payload, JWT_SECRET, algorithm="HS256")


def verificar_token(request: Request):
    auth = request.headers.get("Authorization", "")
    if not auth.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="No autorizado")
    try:
        return jwt.decode(auth.split(" ")[1], JWT_SECRET, algorithms=["HS256"])
    except Exception:
        raise HTTPException(status_code=401, detail="Token inválido")

def solo_admin(payload: dict = Depends(verificar_token)):
    if payload.get("rol") != "administrador":
        raise HTTPException(status_code=403, detail="Solo administradores")
    return payload


# ── Auth ──────────────────────────────────────────────────────────────────────
@app.post("/login")
def login(data: LoginRequest):
    usuario = obtener_usuario_sqlite(data.usuario)
    
    if not usuario:
        raise HTTPException(401, "Usuario o contraseña incorrectos")
    
    try:
        ok = bcrypt.checkpw(data.password.encode(), usuario["password"].encode())
    except Exception:
        ok = False
    
    if not ok:
        raise HTTPException(401, "Usuario o contraseña incorrectos")
    
    if usuario["situacion"] != "activo":
        raise HTTPException(403, "Usuario inactivo. Contacte al administrador.")
    
    token = crear_token_dinamico(usuario, data.password)
    
    return {
        "token": token,
        "usuario": usuario["usuario"],
        "nombre": usuario["nombre_usuario"],
        "rol": usuario["rol"],
        "base_datos": usuario["base_datos"],
        "servidor_bd": usuario["servidor_bd"]
    }


# ── Admin Usuarios ────────────────────────────────────────────────────────────
@app.get("/admin/usuarios")
def listar_usuarios(payload: dict = Depends(solo_admin)):
    """Lista usuarios desde SQLite"""
    try:
        conn = sqlite3.connect('/app/usuarios.db')
        cursor = conn.cursor()
        cursor.execute("""
            SELECT usuario, nombre_usuario, servidor_bd, base_datos, api_url, rol, situacion, usuario_bd
            FROM usuarios
        """)
        usuarios = []
        for row in cursor.fetchall():
            usuarios.append({
                "usuario": row[0],
                "nombre_usuario": row[1],
                "servidor_bd": row[2],
                "base_datos": row[3],
                "api_url": row[4],
                "rol": row[5],
                "situacion": row[6],
                "usuario_bd": row[7] or ""
            })
        conn.close()
        return usuarios
    except Exception as e:
        return {"error": str(e)}

@app.post("/admin/usuarios")
def crear_usuario(data: UsuarioRequest, payload: dict = Depends(solo_admin)):
    if data.situacion not in ("activo","inactivo"): raise HTTPException(400, "Situacion invalida")
    conn = sqlite3.connect('/app/usuarios.db')
    c = conn.cursor()
    c.execute("SELECT usuario FROM usuarios WHERE usuario=?", (data.usuario,))
    if c.fetchone():
        conn.close()
        raise HTTPException(400, "El usuario ya existe")
    pwd_hash = bcrypt.hashpw(data.password.encode(), bcrypt.gensalt()).decode()
    c.execute("INSERT INTO usuarios (usuario,nombre_usuario,password,servidor_bd,base_datos,api_url,rol,situacion,usuario_bd,password_bd) VALUES (?,?,?,?,?,?,?,?,?,?)",
              (data.usuario, data.nombre_usuario, pwd_hash, data.servidor_bd, data.base_datos, data.api_url, data.rol, data.situacion, data.usuario_bd, data.password_bd))
    conn.commit()
    conn.close()
    return {"ok": True}

@app.put("/admin/usuarios/{usuario}")
def editar_usuario(usuario: str, data: UsuarioRequest, payload: dict = Depends(solo_admin)):
    if data.situacion not in ("activo","inactivo"): raise HTTPException(400, "Situacion invalida")
    conn = sqlite3.connect('/app/usuarios.db')
    c = conn.cursor()
    c.execute("SELECT usuario FROM usuarios WHERE usuario=?", (usuario,))
    if not c.fetchone():
        conn.close()
        raise HTTPException(404, "Usuario no encontrado")
    c.execute("UPDATE usuarios SET nombre_usuario=?,servidor_bd=?,base_datos=?,api_url=?,rol=?,situacion=?,usuario_bd=? WHERE usuario=?",
              (data.nombre_usuario, data.servidor_bd, data.base_datos, data.api_url, data.rol, data.situacion, data.usuario_bd, usuario))
    if data.password:
        pwd_hash = bcrypt.hashpw(data.password.encode(), bcrypt.gensalt()).decode()
        c.execute("UPDATE usuarios SET password=? WHERE usuario=?", (pwd_hash, usuario))
    if data.password_bd:
        c.execute("UPDATE usuarios SET password_bd=? WHERE usuario=?", (data.password_bd, usuario))
    conn.commit()
    conn.close()
    return {"ok": True}

@app.delete("/admin/usuarios/{usuario}")
def eliminar_usuario(usuario: str, payload: dict = Depends(solo_admin)):
    if usuario == payload["sub"]: raise HTTPException(400, "No puedes eliminarte a ti mismo")
    conn = sqlite3.connect('/app/usuarios.db')
    c = conn.cursor()
    c.execute("DELETE FROM usuarios WHERE usuario=?", (usuario,))
    conn.commit()
    conn.close()
    return {"ok": True}


# ── PDF Azure ─────────────────────────────────────────────────────────────────
@app.get("/pdf/{tabla}/{internal_id}")
async def obtener_pdf(tabla: str, internal_id: int, payload: dict = Depends(verificar_token)):
    nombre_tabla = tabla.replace("CAPLAN_", "")
    if nombre_tabla not in TABLAS_CON_DOCUMENTO:
        raise HTTPException(400, "Tabla no soporta descarga de PDF")
    api_url_usuario = payload.get("api_url", "https://hurpayai.ximery.com/api")
    if False:
        raise HTTPException(400, "Tabla no soporta descarga de PDF")
    try:
        token_azure = await obtener_token_azure(payload["sub"], payload.get("pwd_azure", ""), api_url=api_url_usuario)
        pdf_bytes   = await descargar_pdf_azure(nombre_tabla, internal_id, token_azure, api_url=api_url_usuario)
    except httpx.HTTPStatusError as e:
        raise HTTPException(502, f"Error en Azure: {e.response.status_code}")
    except Exception as e:
        raise HTTPException(502, f"Error al conectar con Azure: {str(e)}")
    return StreamingResponse(io.BytesIO(pdf_bytes), media_type="application/pdf",
        headers={"Content-Disposition": f"inline; filename={nombre_tabla}_{internal_id}.pdf"})


# ── Re-indexar (admin) ────────────────────────────────────────────────────────
@app.post("/admin/reindexar")
def reindexar(payload: dict = Depends(solo_admin)):
    """Recarga el Excel y re-indexa ChromaDB. Llamar tras actualizar el Excel."""
    global MODELO, coleccion_modelo
    MODELO = cargar_todo()
    try:
        coleccion_modelo.delete(ids=["__control__"])
    except Exception:
        pass
    coleccion_modelo = indexar_modelo_en_chroma()

# ── Indexar PDFs de manuales en ChromaDB ─────────────────────────────────────
    return {
        "ok":                   True,
        "docs_indexados":       coleccion_modelo.count(),
        "hash_excel":           MODELO["hash_excel"],
        "tablas_mapeadas":      len(MODELO["eq_tablas"]),
        "columnas_mapeadas":    len(MODELO["eq_cols"]),
        "relaciones_resueltas": len(MODELO["relaciones_resueltas"]),
        "reglas_cargadas":      len([l for l in MODELO.get("reglas", "").splitlines() if l.strip()]),
    }


# ── Chat ──────────────────────────────────────────────────────────────────────
# Rate limit: máximo 30 consultas por minuto por IP

def detectar_complejidad_consulta(pregunta: str) -> str:
    """
    Manual -> Haiku | SQL/Datos -> Sonnet (precision 100%)
    """
    pregunta_lower = pregunta.lower()
    palabras_manual = [
        'como', 'cómo', 'que es', 'qué es', 'explica', 'explícame',
        'pasos', 'procedimiento', 'instrucciones',
        'manual', 'guía', 'ayuda', 'puedo', 'debo', 'tengo que',
        'funciona', 'significa', 'definición', 'concepto',
        'ingresar', 'ingreso', 'crear', 'creo', 'registrar', 'registro',
        'emitir', 'emisión', 'firmar', 'firma', 'configurar', 'configuración',
        'generar', 'enviar', 'proceso de', 'cómo se', 'como se',
        'qué pasos', 'que pasos', 'cómo puedo', 'como puedo',
        'eliminar', 'elimino', 'anular', 'anulo', 'editar', 'edito',
        'modificar', 'modifico', 'borrar', 'borro', 'actualizar', 'actualizo',
        'cambiar', 'cambio'
    ]
    if any(p in pregunta_lower for p in palabras_manual):
        return 'manual'
    return 'sql'

def seleccionar_modelo(complejidad: str) -> str:
    """
    Manual -> Haiku | SQL -> Sonnet (precision 100%)
    """
    if complejidad == 'manual':
        return "claude-haiku-4-5-20251001"
    else:
        return "claude-sonnet-4-5-20250929"


@app.post("/chat")
@limiter.limit("30/minute")
def chat(data: ChatRequest, request: Request, payload: dict = Depends(verificar_token)):
    servidor_bd_usuario = payload.get("servidor_bd", SQL_SERVER)
    base_datos_usuario  = payload.get("base_datos",  SQL_DATABASE)
    api_url_usuario     = payload.get("api_url",     "https://hurpayai.ximery.com/api")
    usuario_bd_usuario  = payload.get("usuario_bd",  SQL_USER)
    password_bd_usuario = payload.get("password_bd", SQL_PASSWORD)
    usuario_id          = payload.get("sub", "anonimo")
    
    # ── 1. Caché de consultas ────────────────────────────────────────────────
    cache_key = _cache_key(usuario_id, data.mensaje)
    respuesta_cacheada = _cache_get(cache_key)
    if respuesta_cacheada:
        respuesta_cacheada["_cache"] = True
        return respuesta_cacheada
    
    # ── 2. Manuales: buscar fragmentos relevantes en ChromaDB (RAG) ──────────
    try:
        resultados = coleccion.query(query_texts=[data.mensaje], n_results=3)
        contexto_manuales = ""
        if resultados["documents"][0]:
            contexto_manuales = "\n\n".join([
                f"Fuente: {resultados['metadatas'][0][i].get('fuente','')}, "
                f"Pagina: {resultados['metadatas'][0][i].get('pagina','')}\n{doc}"
                for i, doc in enumerate(resultados["documents"][0])
            ])
    except Exception as e:
        contexto_manuales = f"Error buscando manuales: {str(e)}"
    
    # ── 3. Historial ─────────────────────────────────────────────────────────
    historial_ai = []
    if data.historial:
        historial_ai = [{"role": m.get("rol", "user"), "content": m.get("contenido", "")} 
                       for m in data.historial[-6:]]
    historial_ai.append({"role": "user", "content": data.mensaje})
    
    # ── 4. Detectar complejidad y seleccionar modelo ──────────────────────────
    complejidad = detectar_complejidad_consulta(data.mensaje)
    modelo_seleccionado = seleccionar_modelo(complejidad)
    
    # ── 5. System prompt con Prompt Caching ──────────────────────────────────
    bloque_estatico = f"""Eres el asistente virtual de {NOMBRE_EMPRESA}, especializado en la aplicación de nóminas Hurpay para Perú.
Tienes DOS fuentes de información:
1. MANUALES (PDFs indexados): para preguntas sobre cómo usar el sistema.
2. BASE DE DATOS SQL Server: para consultas de datos reales de trabajadores, nóminas, contratos, etc.

=== MODELO DE DATOS COMPLETO DE HURPAY ===
{MODELO["prompt_modelo"]}

=== INSTRUCCIONES ESTRICTAS ===
Las reglas del sistema están definidas en CONDICIONES GENERALES arriba. Debes seguirlas todas sin excepción.
Formato OBLIGATORIO para SELECTs:
<SQL>
SELECT ...
</SQL>
Luego explica brevemente qué hace la consulta y qué datos devuelve."""
    
    bloque_dinamico = f"""Usuario activo: Base de datos={base_datos_usuario} | API={api_url_usuario}
=== MANUALES (fragmentos relevantes) ===
{contexto_manuales or "No se encontró información relevante en los manuales."}"""
    
    system_con_cache = [
        {
            "type": "text",
            "text": bloque_estatico,
            "cache_control": {"type": "ephemeral", "ttl": "1h"},
        },
        {
            "type": "text",
            "text": bloque_dinamico,
        },
    ]
    
    # ── 6. Llamar a Claude ───────────────────────────────────────────────────
    try:
        respuesta_claude = cliente_ai.messages.create(
            model=modelo_seleccionado,
            max_tokens=3000,
            system=system_con_cache,
            messages=historial_ai,
            extra_headers={"anthropic-beta": "prompt-caching-2024-07-31"},
        )
        texto_respuesta = respuesta_claude.content[0].text
        guardar_log_consulta(usuario_id, modelo_seleccionado, respuesta_claude.usage, data.mensaje)
    except Exception as e:
        return {"error": f"Error llamando a Claude: {str(e)}"}
    
    # ── 7. Procesar SQL si lo generó ─────────────────────────────────────────
    resultado_sql = None
    tabla_detectada = None
    
    if "<SQL>" in texto_respuesta and "</SQL>" in texto_respuesta:
        inicio = texto_respuesta.index("<SQL>") + 5
        fin = texto_respuesta.index("</SQL>")
        query_generado = texto_respuesta[inicio:fin].strip()
        
        # Validar seguridad
        es_seguro, msg_seguridad = validar_sql_seguro_mejorado(query_generado)
        if not es_seguro:
            resultado_sql = {"error": msg_seguridad}
        else:
            resultado_sql = ejecutar_sql(query_generado, base_datos=base_datos_usuario, servidor=servidor_bd_usuario, usuario_bd=usuario_bd_usuario, password_bd=password_bd_usuario)
            
            # Detectar tabla para PDF
            q_upper = query_generado.upper()
            for tabla in TABLAS_CON_DOCUMENTO:
                if tabla.upper() in q_upper or f"CAPLAN_{tabla}".upper() in q_upper:
                    tabla_detectada = tabla
                    break
        
        texto_respuesta = texto_respuesta.replace(
            f"<SQL>{query_generado}</SQL>", f"```sql\n{query_generado}\n```")
    
    # ── 8. Respuesta final ───────────────────────────────────────────────────
    respuesta_final = {"respuesta": texto_respuesta, "modelo_usado": modelo_seleccionado}
    
    if resultado_sql:
        respuesta_final["datos"] = resultado_sql
        columnas = resultado_sql.get("columnas") or []
        if tabla_detectada and any("InternalId" in c for c in columnas):
            respuesta_final["tabla_origen"] = tabla_detectada
    
    # ── 9. Caché ─────────────────────────────────────────────────────────────
    if not resultado_sql:
        _cache_set(cache_key, respuesta_final)
    
    return respuesta_final
@app.get("/admin/consumo/{usuario}")
def obtener_consumo(usuario: str, fecha_desde: str = None, fecha_hasta: str = None, payload: dict = Depends(solo_admin)):
    try:
        conn = sqlite3.connect('/app/usuarios.db')
        query = "SELECT id, fecha, hora, usuario, modelo, input_tokens, cache_read_tokens, cache_write_tokens_5m, cache_write_tokens_1h, output_tokens, costo_usd, pregunta FROM log_consultas WHERE usuario=?"
        params = [usuario]
        if fecha_desde:
            query += " AND fecha >= ?"
            params.append(fecha_desde)
        if fecha_hasta:
            query += " AND fecha <= ?"
            params.append(fecha_hasta)
        query += " ORDER BY fecha DESC, hora DESC"
        rows = conn.execute(query, params).fetchall()
        conn.close()
        consultas = []
        total_costo = 0
        total_input = 0
        total_output = 0
        for r in rows:
            consultas.append({
                "id": r[0], "fecha": r[1], "hora": r[2], "usuario": r[3],
                "modelo": r[4], "input_tokens": r[5], "cache_read_tokens": r[6],
                "cache_write_tokens_5m": r[7], "cache_write_tokens_1h": r[8],
                "output_tokens": r[9], "costo_usd": r[10], "pregunta": r[11]
            })
            total_costo += r[10] or 0
            total_input += r[5] or 0
            total_output += r[9] or 0
        return {
            "usuario": usuario,
            "consultas": consultas,
            "total_consultas": len(consultas),
            "total_input_tokens": total_input,
            "total_output_tokens": total_output,
            "total_costo_usd": round(total_costo, 6)
        }
    except Exception as e:
        raise HTTPException(500, f"Error: {str(e)}")



@app.get("/health")
def health():
    prompt_tokens = len(MODELO.get("prompt_modelo", "")) // 4
    return {
        "status":               "ok",
        "empresa":              NOMBRE_EMPRESA,
        "hash_excel":           MODELO.get("hash_excel", ""),
        "docs_indexados":       coleccion_modelo.count(),
        "tablas_mapeadas":      len(MODELO["eq_tablas"]),
        "columnas_mapeadas":    len(MODELO["eq_cols"]),
        "relaciones_resueltas": len(MODELO["relaciones_resueltas"]),
        "prompt_tokens_aprox":  prompt_tokens,
        "cache_consultas":      len(_CACHE_CONSULTAS),
        "cors_origins":         ALLOWED_ORIGINS,
    }

@app.get("/test-sql")
def test_sql(payload: dict = Depends(verificar_token)):
    try:
        servidor_bd_usuario = payload.get("servidor_bd", SQL_SERVER)
        base_datos_usuario = payload.get("base_datos", SQL_DATABASE)
        conn = get_conexion_sql(base_datos_usuario, servidor_bd_usuario)
        conn.cursor().execute("SELECT TOP 1 1 AS test")
        conn.close()
        return {"status": "conectado", "servidor": servidor_bd_usuario, "base_datos": base_datos_usuario}
    except Exception as e:
        return {"error": str(e)}


# ── Admin: limpiar caché de consultas ────────────────────────────────────────
@app.post("/admin/cache/limpiar")
def limpiar_cache(payload: dict = Depends(solo_admin)):
    """Vacía el caché de consultas en memoria. Útil tras actualizar manuales o el Excel."""
    cantidad = len(_CACHE_CONSULTAS)
    _CACHE_CONSULTAS.clear()
    return {"ok": True, "entradas_eliminadas": cantidad}



# Prompts Favoritos
@app.get("/admin/get-categorias")
def listar_categorias_usuario(usuario: str, payload: dict = Depends(solo_admin)):
    conn = sqlite3.connect("/app/usuarios.db")
    rows = conn.execute("SELECT id, nombre_categoria, fecha FROM prompt_categorias WHERE usuario=? ORDER BY nombre_categoria", (usuario,)).fetchall()
    conn.close()
    return {"categorias": [{"id": r[0], "nombre_categoria": r[1], "fecha": r[2]} for r in rows]}

@app.post("/admin/add-categoria")
def crear_categoria(usuario: str, data: CategoriaRequest, payload: dict = Depends(solo_admin)):
    conn = sqlite3.connect("/app/usuarios.db")
    conn.execute("INSERT INTO prompt_categorias (usuario, nombre_categoria) VALUES (?, ?)", (usuario, data.nombre_categoria))
    conn.commit()
    conn.close()
    return {"ok": True, "mensaje": "Categoria creada para " + usuario}

@app.delete("/admin/del-categoria/{id}")
def eliminar_categoria(id: int, payload: dict = Depends(solo_admin)):
    conn = sqlite3.connect("/app/usuarios.db")
    total = conn.execute("SELECT COUNT(*) FROM prompts_favoritos WHERE categoria_id=?", (id,)).fetchone()[0]
    if total > 0:
        conn.close()
        raise HTTPException(400, "No se puede eliminar la categoria porque tiene " + str(total) + " prompts asociados. Elimina los prompts primero.")
    conn.execute("DELETE FROM prompt_categorias WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return {"ok": True}

@app.get("/prompts/categorias")
def listar_mis_categorias(payload: dict = Depends(verificar_token)):
    usuario = payload.get("sub")
    conn = sqlite3.connect("/app/usuarios.db")
    rows = conn.execute("SELECT id, nombre_categoria FROM prompt_categorias WHERE usuario=? ORDER BY nombre_categoria", (usuario,)).fetchall()
    conn.close()
    return {"categorias": [{"id": r[0], "nombre_categoria": r[1]} for r in rows]}

@app.get("/prompts")
def listar_prompts(categoria_id: int = None, nombre: str = None, payload: dict = Depends(verificar_token)):
    usuario = payload.get("sub")
    query = "SELECT pf.id, pf.nombre, pf.prompt, pf.fecha, pc.nombre_categoria, pf.categoria_id FROM prompts_favoritos pf JOIN prompt_categorias pc ON pc.id=pf.categoria_id WHERE pf.usuario=?"
    params = [usuario]
    if categoria_id:
        query += " AND pf.categoria_id=?"
        params.append(categoria_id)
    if nombre:
        query += " AND pf.nombre LIKE ?"
        params.append("%" + nombre + "%")
    query += " ORDER BY pc.nombre_categoria, pf.nombre"
    conn = sqlite3.connect("/app/usuarios.db")
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return {"prompts": [{"id": r[0], "nombre": r[1], "prompt": r[2], "fecha": r[3], "categoria": r[4], "categoria_id": r[5]} for r in rows], "total": len(rows)}

@app.post("/prompts")
def guardar_prompt(data: PromptFavoritoRequest, payload: dict = Depends(verificar_token)):
    usuario = payload.get("sub")
    conn = sqlite3.connect("/app/usuarios.db")
    conn.execute("INSERT INTO prompts_favoritos (usuario, categoria_id, nombre, prompt) VALUES (?, ?, ?, ?)", (usuario, data.categoria_id, data.nombre, data.prompt))
    conn.commit()
    conn.close()
    return {"ok": True, "mensaje": "Prompt guardado correctamente"}

@app.put("/prompts/{id}")
def editar_prompt(id: int, data: PromptFavoritoRequest, payload: dict = Depends(verificar_token)):
    usuario = payload.get("sub")
    conn = sqlite3.connect("/app/usuarios.db")
    existing = conn.execute("SELECT id FROM prompts_favoritos WHERE id=? AND usuario=?", (id, usuario)).fetchone()
    if not existing:
        conn.close()
        raise HTTPException(404, "Prompt no encontrado")
    conn.execute("UPDATE prompts_favoritos SET categoria_id=?, nombre=?, prompt=? WHERE id=? AND usuario=?", (data.categoria_id, data.nombre, data.prompt, id, usuario))
    conn.commit()
    conn.close()
    return {"ok": True, "mensaje": "Prompt actualizado correctamente"}

@app.delete("/prompts/{id}")
def eliminar_prompt(id: int, payload: dict = Depends(verificar_token)):
    usuario = payload.get("sub")
    conn = sqlite3.connect("/app/usuarios.db")
    existing = conn.execute("SELECT id FROM prompts_favoritos WHERE id=? AND usuario=?", (id, usuario)).fetchone()
    if not existing:
        conn.close()
        raise HTTPException(404, "Prompt no encontrado")
    conn.execute("DELETE FROM prompts_favoritos WHERE id=? AND usuario=?", (id, usuario))
    conn.commit()
    conn.close()
    return {"ok": True, "mensaje": "Prompt eliminado correctamente"}

app.mount("/", StaticFiles(directory="/app/static", html=True), name="static")

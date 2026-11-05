# -*- coding: utf-8 -*-
"""
Azimut Energía SAS — Suite Cruz Verde
Propiedad intelectual de Azimut Energía SAS.
Prohibida su reproducción, distribución o uso no autorizado.

Módulos:
  1. Integración          — carga archivo Azimut → hoja Traza
  2. Generación CXP       — extrae PSE/Cheque del día, genera plantilla CXP,
                            marca filas como Ejecutado en el Traza
"""

import os
import sys
import json
import re
import unicodedata
import threading
import traceback
import tempfile
import time
from datetime import datetime, date
from decimal import Decimal, InvalidOperation

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
import webview

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

APP_TITLE = "Azimut · Suite Cruz Verde"
CONFIG_FILENAME = "config_traza_cruzverde.json"
CXP_TEMPLATE_NAME = "Plantilla_CXP.xlsx"

CXP_METODOS_DISPONIBLES = [
    "Recobro",
    "PSE",
    "Pendiente Davivienda",
    "Pendiente Bogota",
    "Nuevo PSE",
    "Davivienda",
    "Cuenta Padre",
    "Cheque",
    "Banco de Bogota",
    "Administracion",
]
CXP_FLEX_LARGO_ESPERADO = 52
CXP_COL_EJECUTADO_LETRA = "BT"
CXP_VALOR_EJECUTADO = "Ejecutado"

CXP_CAMPOS_PLANTILLA = [
    ("CUENTA CONTRATO",     1),
    ("CECO R11",            2),
    ("NIT",                 3),
    ("MEDIO DE PAGO",       4),
    ("NUMERO DE FACTURA",   5),
    ("VALOR ANTES DE IVA",  6),
    ("IMPUESTO",            7),
    # columna 8 = SECUENCIA → siempre en blanco
    ("FECHA DE VENCIMIENTO", 9),
    ("FECHA DE EMISION",    10),
    ("OBSERVACION ORACLE",  11),
    ("FLEX",                12),
    # columna 13 = LARGO  → =+LEN(Ln)
    # columna 14 = GRUPO  → "SERVICIOS PUBLICOS"
]
CXP_CAMPOS_FILTRO = ["METODO DE PAGO", "FECHA DE GESTION"]
CXP_CAMPOS_DEDUP  = ["NUMERO DE FACTURA"]

# ---------------------------------------------------------------------------
# Módulo 3: Contabilización — constantes
#
# Tomadas de la hoja "Platilla contabilidad" del archivo real
# (Elaboracion_plantilla_contabilidad.xlsx). Si alguna cambia, se ajusta
# solo aquí.
# ---------------------------------------------------------------------------
WEBADI_FLEX_LARGO_ESPERADO = 52

# "Información estática no cambia" en la plantilla real
WEBADI_EMPRESA = 88
WEBADI_CTA_CONTABLE_GASTO = "511060085"   # cuenta de gasto para servicios públicos
WEBADI_TIPO_CUENTA = ".00.00000."          # bloque fijo que sigue a la línea de negocio
WEBADI_PUNTO_COMPRA = "99999"
WEBADI_LINEA_NEGOCIO = "999"
WEBADI_UEN = None  # columna "UEN" de Plantilla contabilidad: siempre vacía en el archivo real

# Columnas de "Plantilla contabilidad" en el archivo real (1-indexadas).
# Banco, Cuenta bancaria, Número de Trx, Fecha de Trx, Importe y Observación
# Oracle (columnas 8-12 y 17) NO se escriben desde aquí: se completan a mano
# cuando se sabe con qué banco se pagó.
PC_SHEET_ALIASES = ("Platilla contabilidad", "Plantilla contabilidad")  # el archivo real trae la hoja sin la "n"
PC_COL = {
    "FECHA": 1, "CONCEPTO": 2, "CODIGO_CLIENTE": 3, "CCOSTO": 4, "LOCAL": 5,
    "NIT": 6, "PROVEEDOR": 7, "BANCO": 8, "CUENTA_BANCARIA": 9, "NUMERO_TRX": 10,
    "FECHA_TRX": 11, "IMPORTE": 12, "FECHA_FACTURA": 13, "FACTURA": 14,
    "GASTO_A_CONTABILIZAR": 15, "IVA": 16, "OBSERVACION_ORACLE": 17, "EMPRESA": 18,
    "UEN": 19, "CC_R11": 20, "TIPO_CUENTA": 21, "USO_FUTURO": 22, "CTA_CONTABLE": 23,
    "SUC": 24, "CC_R12": 25, "PV": 26, "PC": 27, "LN": 28, "CONCATENAR_FLEX": 29,
    "LARGO": 30, "BANCO2": 33, "CUENTA_CONTABLE_BANCO": 34, "NIT_BANCO": 35,
    "PROVEEDOR_SIN_DIGITO": 36,
}

ALIAS_GROUPS = [
    {"PERIODICIDAD", "FRECUENCIA"},
    {"FACTURA", "NUMERO DE FACTURA"},
    {"PERIODO INICIAL - FECHA", "PERIODO INICIAL"},
    {"PERIODO FINAL - FECHA",   "PERIODO FINAL"},
    {"VALOR IMPUESTO - IVA", "IMPUESTO"},
    {"CONSUMO KVR ENERGIA REACTIVA", "ENERGIA REACTIVA"},
]

HEADER_KEYWORDS = {
    "CIUDAD", "EMPRESA", "FACTURA", "DIRECCION", "RESPONSABLE", "CONSUMO",
    "VALOR", "PERIODICIDAD", "FRECUENCIA", "NIT", "SERVICIO", "CODIGO",
    "FECHA", "FLEX", "AREA", "FORMATO",
}

# ---------------------------------------------------------------------------
# Helpers de rutas
# ---------------------------------------------------------------------------

def app_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def gui_path(relative):
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "gui", relative)

def template_path(relative):
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "templates", relative)

# ---------------------------------------------------------------------------
# Config (últimas rutas por PC, no se comparte entre equipos)
# ---------------------------------------------------------------------------

def load_config():
    path = os.path.join(app_base_dir(), CONFIG_FILENAME)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_config(data):
    path = os.path.join(app_base_dir(), CONFIG_FILENAME)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ---------------------------------------------------------------------------
# Guardado atómico — protege el archivo destino si se cierra la app a medias
# ---------------------------------------------------------------------------

def guardar_atomico(workbook, destino):
    """Guarda en un temporal al lado del destino y solo reemplaza cuando el
    guardado terminó 100%. Si la app se cierra a mitad, el archivo original
    queda intacto.

    Reintenta varias veces antes de rendirse: en carpetas de OneDrive/
    SharePoint sincronizadas es normal que, por una fracción de segundo,
    otro proceso (el propio sincronizador, un antivirus) tenga el archivo
    tomado sin que haya ningún Excel abierto de verdad."""
    carpeta = os.path.dirname(os.path.abspath(destino)) or "."

    # Si Excel (de cualquiera, incluida otra persona en un archivo
    # compartido) tiene el archivo abierto, deja este archivo oculto al
    # lado — es la señal más confiable de que SÍ está abierto en algún lado.
    nombre_lock = os.path.join(carpeta, "~$" + os.path.basename(destino))
    if os.path.exists(nombre_lock):
        raise PermissionError(
            f"El archivo está abierto en Excel (se encontró '{os.path.basename(nombre_lock)}' "
            "en la misma carpeta). Ciérralo — puede estar abierto en otro equipo si la carpeta "
            "se sincroniza por OneDrive/SharePoint — e intenta de nuevo. No se modificó."
        )

    intentos = 5
    espera = 0.4
    ultimo_error = None
    for intento in range(1, intentos + 1):
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", prefix="~tmp_azimut_", dir=carpeta)
        os.close(fd)
        try:
            workbook.save(tmp)
            os.replace(tmp, destino)
            return
        except OSError as e:
            ultimo_error = e
            try:
                os.remove(tmp)
            except OSError:
                pass
            if intento < intentos:
                time.sleep(espera)
                espera *= 2
                continue
            raise PermissionError(
                f"No se pudo guardar tras {intentos} intentos. Error del sistema: {ultimo_error}. "
                "Si la carpeta se sincroniza con OneDrive/SharePoint, espera a que termine de "
                "sincronizar (ícono de la nube) e inténtalo de nuevo; si el problema sigue, "
                "revisa que no tengas el archivo abierto en otro equipo ni marcado como solo lectura."
            ) from e
        except Exception:
            try:
                os.remove(tmp)
            except OSError:
                pass
            raise

# ---------------------------------------------------------------------------
# Helpers compartidos de Excel
# ---------------------------------------------------------------------------

def normalize(value):
    if value is None:
        return ""
    s = str(value).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip()

def to_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def detect_header_row(ws, max_scan=15):
    best_row, best_score = 1, -1
    max_scan = min(max_scan, ws.max_row or 1)
    for r in range(1, max_scan + 1):
        score = 0
        for c in range(1, (ws.max_column or 1) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            nv = normalize(v)
            for kw in HEADER_KEYWORDS:
                if kw in nv:
                    score += 1
                    break
        if score > best_score:
            best_score, best_row = score, r
    return best_row if best_score >= 3 else 1

def read_headers(ws, header_row):
    headers = []
    for c in range(1, (ws.max_column or 1) + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is not None and str(v).strip():
            headers.append((c, v, normalize(v)))
    return headers

def find_col_for_group(headers_dict, names):
    for name in names:
        if name in headers_dict:
            return headers_dict[name]
    return None

def build_mapping(source_headers, dest_headers_dict):
    mapping, unmatched = {}, []
    for col_idx, raw, norm in source_headers:
        if norm in dest_headers_dict:
            mapping[col_idx] = dest_headers_dict[norm]
            continue
        matched = False
        for group in ALIAS_GROUPS:
            if norm in group:
                for alt in group:
                    if alt in dest_headers_dict:
                        mapping[col_idx] = dest_headers_dict[alt]
                        matched = True
                        break
            if matched:
                break
        if not matched:
            unmatched.append(raw)
    return mapping, unmatched

def find_last_row_with_data(ws, header_row, dest_headers_dict):
    anchor_col = (
        dest_headers_dict.get("EMPRESA")
        or dest_headers_dict.get("NOMBRE SUCURSAL")
        or (min(dest_headers_dict.values()) if dest_headers_dict else 1)
    )
    last_row = header_row
    r = header_row + 1
    while True:
        v = ws.cell(row=r, column=anchor_col).value
        if v is None or str(v).strip() == "":
            break
        last_row = r
        r += 1
    return last_row

# ---------------------------------------------------------------------------
# Módulo 1: Integración (copia Azimut → hoja Traza)
# ---------------------------------------------------------------------------

def copy_azimut_to_traza(traza_path, hoja_destino, azimut_path, log):
    """Copia Azimut → Traza usando lectura por filas para reducir accesos celda a celda."""
    log("Abriendo archivo Azimut (origen)...")
    wb_src = openpyxl.load_workbook(azimut_path, data_only=True, read_only=True)
    ws_src = wb_src[wb_src.sheetnames[0]]

    src_header_row = detect_header_row(ws_src)
    source_headers = read_headers(ws_src, src_header_row)
    log(f"Encabezados de origen detectados en la fila {src_header_row} "
        f"({len(source_headers)} columnas).")

    log("Abriendo archivo Traza (destino)...")
    wb_dest = openpyxl.load_workbook(traza_path)
    if hoja_destino not in wb_dest.sheetnames:
        wb_src.close()
        raise ValueError(f"La hoja '{hoja_destino}' ya no existe en el archivo Traza.")
    ws_dest = wb_dest[hoja_destino]

    dest_header_row = detect_header_row(ws_dest)
    dest_headers_list = read_headers(ws_dest, dest_header_row)
    dest_headers_dict = {norm: col for col, raw, norm in dest_headers_list}
    log(f"Encabezados de destino detectados en la fila {dest_header_row} "
        f"({len(dest_headers_list)} columnas).")

    mapping, unmatched = build_mapping(source_headers, dest_headers_dict)
    log(f"Columnas con match: {len(mapping)} de {len(source_headers)}.")
    if unmatched:
        log("Columnas del archivo Azimut sin match (se ignoran): " + ", ".join(unmatched))

    last_row = find_last_row_with_data(ws_dest, dest_header_row, dest_headers_dict)
    log(f"Última fila con datos en destino: {last_row}. Se escribe desde la fila {last_row + 1}.")
    template_row = last_row if last_row > dest_header_row else None

    source_headers_dict = {norm: col for col, raw, norm in source_headers}
    key_src_cols = [
        find_col_for_group(source_headers_dict, {"CUENTA CONTRATO"}),
        find_col_for_group(source_headers_dict, {"FACTURA", "NUMERO DE FACTURA"}),
        find_col_for_group(source_headers_dict, {"SERVICIO"}),
    ]
    key_dest_cols = [
        find_col_for_group(dest_headers_dict, {"CUENTA CONTRATO"}),
        find_col_for_group(dest_headers_dict, {"FACTURA", "NUMERO DE FACTURA"}),
        find_col_for_group(dest_headers_dict, {"SERVICIO"}),
    ]
    dedup_enabled = all(c is not None for c in key_src_cols + key_dest_cols)

    existing_keys = set()
    if dedup_enabled and last_row > dest_header_row:
        min_c, max_c = min(key_dest_cols), max(key_dest_cols)
        rel = [c - min_c for c in key_dest_cols]
        for vals in ws_dest.iter_rows(
            min_row=dest_header_row + 1, max_row=last_row,
            min_col=min_c, max_col=max_c, values_only=True
        ):
            key = tuple(normalize(vals[i]) for i in rel)
            if key[1]:
                existing_keys.add(key)
        log(f"Verificación de duplicados activa. Registros existentes: {len(existing_keys)}.")
    else:
        log("Verificación de duplicados no disponible (faltan columnas clave)." if not dedup_enabled
            else "Verificación de duplicados activa. Registros existentes: 0.")

    # Preparar formatos una sola vez, en vez de consultarlos para cada celda escrita.
    template_formats = {}
    if template_row:
        for dest_col in set(mapping.values()):
            template_formats[dest_col] = ws_dest.cell(row=template_row, column=dest_col).number_format

    rows_written = 0
    duplicates_skipped = 0
    write_r = last_row + 1
    src_max_col = ws_src.max_column or 1
    mapped_src = [(src_col, dest_col) for src_col, dest_col in mapping.items()]
    src_header_cols = [col_idx for col_idx, _, _ in source_headers]

    # Una sola lectura secuencial del archivo Azimut.
    for values in ws_src.iter_rows(
        min_row=src_header_row + 1,
        max_row=ws_src.max_row,
        min_col=1,
        max_col=src_max_col,
        values_only=True,
    ):
        if not any(values[c - 1] not in (None, "") for c in src_header_cols):
            continue

        if dedup_enabled:
            key_vals = tuple(normalize(values[c - 1]) for c in key_src_cols)
            if key_vals[1]:
                if key_vals in existing_keys:
                    duplicates_skipped += 1
                    continue
                existing_keys.add(key_vals)

        for src_col, dest_col in mapped_src:
            target = ws_dest.cell(row=write_r, column=dest_col)
            target.value = values[src_col - 1]
            fmt = template_formats.get(dest_col)
            if fmt:
                target.number_format = fmt

        rows_written += 1
        write_r += 1

    wb_src.close()
    log(f"Filas nuevas a guardar: {rows_written}.")
    if duplicates_skipped:
        log(f"Registros duplicados omitidos: {duplicates_skipped}.")
    log("Guardando archivo Traza...")

    try:
        guardar_atomico(wb_dest, traza_path)
    except PermissionError:
        raise PermissionError(
            "El archivo Traza está abierto o en uso. Ciérralo e intenta de nuevo. "
            "No se modificó el archivo."
        )

    return {
        "rows_written": rows_written,
        "duplicates_skipped": duplicates_skipped,
        "dedup_enabled": dedup_enabled,
        "unmatched": unmatched,
        "matched_cols": len(mapping),
    }

# ---------------------------------------------------------------------------
# Módulo 2: Generación Plantillas CXP
# ---------------------------------------------------------------------------

def _to_decimal(value):
    """Convierte valores monetarios comunes de Excel/texto a Decimal."""
    if value in (None, ""):
        return Decimal("0")
    if isinstance(value, bool):
        return Decimal(int(value))
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))

    s = str(value).strip().replace("$", "").replace(" ", "")
    if not s:
        return Decimal("0")

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        partes = s.split(",")
        if len(partes[-1]) in (1, 2):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "." in s:
        partes = s.split(".")
        if len(partes) > 2 or (len(partes) == 2 and len(partes[-1]) == 3):
            s = s.replace(".", "")

    try:
        return Decimal(s)
    except InvalidOperation:
        raise ValueError(f"No pude interpretar el valor monetario '{value}'.")


def _codigo_sin_ceros(value):
    """Normaliza códigos como 06398, 6398 o 6398.0 al mismo valor lógico."""
    if value in (None, ""):
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".", 1)[0]
    solo_digitos = re.sub(r"\D", "", s)
    if not solo_digitos:
        return normalize(value)
    return solo_digitos.lstrip("0") or "0"


def _analizar_cxp(traza_path, hoja, metodos_filtro, log):
    """Analiza la Traza una sola vez y deja listo todo lo necesario para generar el CXP."""
    metodos_norm = {normalize(m) for m in metodos_filtro}

    log("Abriendo archivo Traza...")
    wb_src = openpyxl.load_workbook(traza_path, data_only=True, read_only=True)
    try:
        if hoja not in wb_src.sheetnames:
            raise ValueError(f"La hoja '{hoja}' ya no existe en el Traza seleccionado.")
        ws = wb_src[hoja]

        header_row = detect_header_row(ws)
        headers = read_headers(ws, header_row)
        headers_dict = {norm: col for col, raw, norm in headers}

        nombres_necesarios = list(dict.fromkeys(
            [n for n, _ in CXP_CAMPOS_PLANTILLA] + CXP_CAMPOS_FILTRO + ["SERVICIO", "PUNTO DE VENTA"]
        ))
        cols, faltantes = {}, []
        for nombre in nombres_necesarios:
            col = headers_dict.get(normalize(nombre))
            cols[nombre] = col
            if col is None:
                faltantes.append(nombre)
        if faltantes:
            raise ValueError("No encontré estas columnas: " + ", ".join(faltantes))

        col_ejecutado = column_index_from_string(CXP_COL_EJECUTADO_LETRA)
        max_col = ws.max_column or 0
        if max_col < col_ejecutado:
            raise ValueError(
                f"La hoja no llega hasta la columna {CXP_COL_EJECUTADO_LETRA} "
                f"(necesaria para marcar/verificar '{CXP_VALOR_EJECUTADO}')."
            )

        hoy = date.today()
        log(f"Filtrando: {metodos_filtro} · fecha {hoy.strftime('%d/%m/%Y')}...")

        # Leemos solo hasta la mayor columna que realmente necesitamos.
        needed_cols = set(cols.values()) | {col_ejecutado}
        min_needed = min(needed_cols)
        max_needed = max(needed_cols)
        idx = {c: c - min_needed for c in needed_cols}

        facturas_ejecutadas = set()
        todas_filas = []
        blancos_consecutivos = 0
        LIMITE_BLANCOS = 300

        # ÚNICO recorrido de la Traza: captura ejecutados, candidatos potenciales y valores para totales.
        for excel_row, values in enumerate(
            ws.iter_rows(
                min_row=header_row + 1,
                max_row=ws.max_row,
                min_col=min_needed,
                max_col=max_needed,
                values_only=True,
            ),
            start=header_row + 1,
        ):
            def val(col):
                return values[idx[col]]

            fila_vacia = all(val(c) in (None, "") for c in cols.values())
            if fila_vacia:
                blancos_consecutivos += 1
                if blancos_consecutivos >= LIMITE_BLANCOS:
                    break
                continue
            blancos_consecutivos = 0

            factura_norm = normalize(val(cols["NUMERO DE FACTURA"]))
            ejecutado = normalize(val(col_ejecutado)) == normalize(CXP_VALOR_EJECUTADO)
            if factura_norm and ejecutado:
                facturas_ejecutadas.add(factura_norm)

            row_data = {nombre: val(cols[nombre]) for nombre, _ in CXP_CAMPOS_PLANTILLA}
            row_data["SERVICIO"] = val(cols["SERVICIO"])
            row_data["PUNTO DE VENTA"] = val(cols["PUNTO DE VENTA"])
            row_data["METODO DE PAGO"] = val(cols["METODO DE PAGO"])
            row_data["FECHA DE GESTION"] = val(cols["FECHA DE GESTION"])
            row_data["_row"] = excel_row
            row_data["_factura_norm"] = factura_norm
            todas_filas.append(row_data)

        candidatos = []
        omitidas_ya_ejecutadas = set()
        for item in todas_filas:
            if normalize(item["METODO DE PAGO"]) not in metodos_norm:
                continue
            fecha_norm = to_date(item["FECHA DE GESTION"])
            if fecha_norm is None or fecha_norm != hoy:
                continue
            fac = item["_factura_norm"]
            if fac and fac in facturas_ejecutadas:
                omitidas_ya_ejecutadas.add(fac)
                continue
            candidatos.append(item)

        if omitidas_ya_ejecutadas:
            log(f"Factura única: se omitieron {len(omitidas_ya_ejecutadas)} factura(s) ya causadas anteriormente.")
            for fac in sorted(omitidas_ya_ejecutadas)[:15]:
                log(f"OMITIDA: Factura {fac} ya fue causada anteriormente.")

        if not candidatos:
            log("No hay registros para enviar hoy. No se generó ningún archivo.")
            return {
                "finales": [], "filas_a_marcar": [], "col_ejecutado": col_ejecutado,
                "duplicates_skipped": len(omitidas_ya_ejecutadas),
                "impuestos_vaciados": 0, "empty": True,
            }

        errores = []
        for item in candidatos:
            flex = "" if item["FLEX"] in (None, "") else str(item["FLEX"]).strip()
            largo = len(flex)
            if largo != CXP_FLEX_LARGO_ESPERADO:
                errores.append(
                    f"Factura {item['NUMERO DE FACTURA']} / Cuenta {item['CUENTA CONTRATO']}: "
                    f"FLEX tiene {largo} caracteres y debe tener {CXP_FLEX_LARGO_ESPERADO}."
                )
                continue
            bloques = flex.split(".")
            if len(bloques) < 5:
                errores.append(
                    f"Factura {item['NUMERO DE FACTURA']} / Cuenta {item['CUENTA CONTRATO']}: "
                    f"no pude obtener el punto de venta del FLEX '{flex}'."
                )
                continue
            punto_flex = _codigo_sin_ceros(bloques[4])
            punto_venta = _codigo_sin_ceros(item["PUNTO DE VENTA"])
            if punto_flex != punto_venta:
                errores.append(
                    f"Factura {item['NUMERO DE FACTURA']} / Cuenta {item['CUENTA CONTRATO']}: "
                    f"Punto de venta del FLEX ({bloques[4]} → {punto_flex}) no coincide "
                    f"con PUNTO DE VENTA ({item['PUNTO DE VENTA']} → {punto_venta})."
                )
        if errores:
            log(f"ERROR: {len(errores)} validación(es) bloqueante(s). El CXP NO se generó.")
            for e in errores[:25]:
                log(f"ERROR: {e}")
            raise ValueError(
                f"No se generó el CXP porque hay {len(errores)} error(es) de validación. "
                "Revisa el detalle mostrado en la consola."
            )

        grupos = {}
        orden_facturas = []
        for item in candidatos:
            key = item["_factura_norm"]
            if not key:
                raise ValueError(
                    f"La fila {item['_row']} no tiene NUMERO DE FACTURA. "
                    "No se generó el CXP porque la factura debe ser única."
                )
            if key not in grupos:
                grupos[key] = []
                orden_facturas.append(key)
            grupos[key].append(item)

        # Totales completos usando las filas ya cargadas en memoria; no se vuelve a recorrer Excel.
        facturas_objetivo = set(orden_facturas)
        total_traza_por_factura = {k: Decimal("0") for k in facturas_objetivo}
        for fila in todas_filas:
            fac = fila["_factura_norm"]
            if fac not in facturas_objetivo:
                continue
            antes = _to_decimal(fila["VALOR ANTES DE IVA"])
            impuesto = _to_decimal(fila["IMPUESTO"])
            servicio = normalize(fila["SERVICIO"])
            impuesto_valido = impuesto if "TELEFON" in servicio else Decimal("0")
            total_traza_por_factura[fac] += antes + impuesto_valido

        finales = []
        filas_a_marcar = []
        impuestos_vaciados = 0
        errores_totales = []
        for key in orden_facturas:
            filas = grupos[key]
            base = dict(filas[0])
            total_antes_iva = Decimal("0")
            total_impuesto = Decimal("0")
            for fila in filas:
                antes = _to_decimal(fila["VALOR ANTES DE IVA"])
                impuesto_original = _to_decimal(fila["IMPUESTO"])
                servicio = normalize(fila["SERVICIO"])
                if "TELEFON" in servicio:
                    impuesto_limpio = impuesto_original
                else:
                    impuesto_limpio = Decimal("0")
                    if impuesto_original != 0:
                        impuestos_vaciados += 1
                        log(
                            f"AJUSTE: Factura {fila['NUMERO DE FACTURA']} / servicio '{fila['SERVICIO']}': "
                            f"se vacía IMPUESTO {impuesto_original} porque solo Telefonía puede llevar impuesto."
                        )
                total_antes_iva += antes
                total_impuesto += impuesto_limpio
                filas_a_marcar.append(fila["_row"])

            base["VALOR ANTES DE IVA"] = float(total_antes_iva)
            base["IMPUESTO"] = float(total_impuesto) if total_impuesto != 0 else None
            total_cxp = total_antes_iva + total_impuesto
            total_traza = total_traza_por_factura[key]
            if total_cxp != total_traza:
                errores_totales.append(
                    f"Factura {base['NUMERO DE FACTURA']}: total a CXP {total_cxp} "
                    f"no coincide con total de todas sus filas en Traza {total_traza}."
                )
            finales.append(base)

        if errores_totales:
            log(f"ERROR: {len(errores_totales)} factura(s) con diferencia de total. El CXP NO se generó.")
            for e in errores_totales[:25]:
                log(f"ERROR: {e}")
            raise ValueError(
                f"No se generó el CXP porque {len(errores_totales)} factura(s) no coinciden "
                "con la suma de todas sus filas en la Traza."
            )

        log(f"Facturas únicas listas para CXP: {len(finales)}.")
        if len(candidatos) > len(finales):
            log(f"Se consolidaron {len(candidatos)} filas en {len(finales)} factura(s) únicas.")
        if impuestos_vaciados:
            log(f"Impuestos eliminados en servicios distintos de Telefonía: {impuestos_vaciados}.")

        return {
            "finales": finales,
            "filas_a_marcar": list(dict.fromkeys(filas_a_marcar)),
            "col_ejecutado": col_ejecutado,
            "duplicates_skipped": len(omitidas_ya_ejecutadas),
            "impuestos_vaciados": impuestos_vaciados,
            "empty": False,
        }
    finally:
        wb_src.close()


def generar_plantilla_cxp(traza_path, hoja, salida_path, metodos_filtro, log, preparado=None):
    """Genera CXP reutilizando la prevalidación para no leer la Traza dos veces."""
    preparado = preparado or _analizar_cxp(traza_path, hoja, metodos_filtro, log)

    if preparado["empty"]:
        return {
            "rows_written": 0,
            "duplicates_skipped": preparado["duplicates_skipped"],
            "alertas_largo": 0,
            "impuestos_vaciados": 0,
            "output_path": None,
            "marcado_ok": True,
            "marcado_error": None,
            "empty": True,
            "_preparado": preparado,
        }

    finales = preparado["finales"]
    filas_a_marcar = preparado["filas_a_marcar"]
    col_ejecutado = preparado["col_ejecutado"]

    if salida_path is None:
        log("Prevalidación correcta. No se detectaron errores bloqueantes.")
        return {
            "rows_written": len(finales),
            "duplicates_skipped": preparado["duplicates_skipped"],
            "alertas_largo": 0,
            "impuestos_vaciados": preparado["impuestos_vaciados"],
            "output_path": None,
            "marcado_ok": True,
            "marcado_error": None,
            "empty": False,
            "_preparado": preparado,
        }

    log(f"Generando archivo con {len(finales)} factura(s) sin volver a leer la Traza...")
    wb_out = openpyxl.load_workbook(template_path(CXP_TEMPLATE_NAME))
    ws_out = wb_out[wb_out.sheetnames[0]]
    formatos = {c: ws_out.cell(row=2, column=c).number_format for c in range(1, 15)}
    formatos[9] = "DD/MM/YYYY"
    formatos[10] = "DD/MM/YYYY"
    grupo_valor = ws_out.cell(row=2, column=14).value or "SERVICIOS PUBLICOS"

    for i, item in enumerate(finales):
        out_row = 2 + i
        for nombre, col in CXP_CAMPOS_PLANTILLA:
            ws_out.cell(row=out_row, column=col, value=item[nombre])
        ws_out.cell(row=out_row, column=13, value=f"=+LEN(L{out_row})")
        ws_out.cell(row=out_row, column=14, value=grupo_valor)
        for c, fmt in formatos.items():
            if fmt:
                ws_out.cell(row=out_row, column=c).number_format = fmt

    try:
        guardar_atomico(wb_out, salida_path)
    except PermissionError:
        raise PermissionError("El archivo de destino está abierto. Ciérralo e intenta de nuevo.")

    marcado_ok = True
    marcado_error = None
    log(f"Marcando {len(filas_a_marcar)} fila(s) como '{CXP_VALOR_EJECUTADO}' en el Traza...")
    try:
        wb_m = openpyxl.load_workbook(traza_path)
        ws_m = wb_m[hoja]
        for fila in filas_a_marcar:
            ws_m.cell(row=fila, column=col_ejecutado, value=CXP_VALOR_EJECUTADO)
        guardar_atomico(wb_m, traza_path)
    except PermissionError:
        marcado_ok = False
        marcado_error = (
            f"El CXP se generó, pero el Traza está abierto y no se pudo marcar '{CXP_VALOR_EJECUTADO}'. "
            f"Ciérralo y vuelve a correr, o márcalas manualmente en la columna {CXP_COL_EJECUTADO_LETRA}."
        )
        log(f"ALERTA: {marcado_error}")
    except Exception as e:
        marcado_ok = False
        marcado_error = f"El CXP se generó, pero falló el marcado en el Traza: {e}"
        log(f"ALERTA: {marcado_error}")

    if marcado_ok:
        log("Listo.")

    return {
        "rows_written": len(finales),
        "duplicates_skipped": preparado["duplicates_skipped"],
        "alertas_largo": 0,
        "impuestos_vaciados": preparado["impuestos_vaciados"],
        "output_path": salida_path,
        "marcado_ok": marcado_ok,
        "marcado_error": marcado_error,
        "empty": False,
    }

# ---------------------------------------------------------------------------
# Módulo 3: Contabilización / Web ADI
#
# A diferencia del CXP (que genera un archivo aparte), este módulo funciona
# como una macro: abre EL MISMO archivo que ya tiene las hojas Traza,
# Centros de Costos, Proveedores, Bancos, Plantilla contabilidad y Web Adi
# (el mismo archivo de siempre), calcula lo que hoy calculan las fórmulas
# de "Plantilla contabilidad", y escribe las filas nuevas directamente en
# "Plantilla contabilidad" y en "Web Adi" de ESE archivo — igual que si
# alguien las hubiera pegado a mano y las fórmulas hubieran corrido.
#
# Lo único que hay que traer de fuera es el detalle de las transacciones
# bancarias (banco, número de transacción, factura que paga, fecha), que
# hoy se escribe a mano leyendo el extracto bancario.
# ---------------------------------------------------------------------------

def _nombre_hoja_pc(wb):
    for nombre in PC_SHEET_ALIASES:
        if nombre in wb.sheetnames:
            return nombre
    return None


# Fila donde ya están las fórmulas listas para "arrastrar" (desplazar) a las
# filas nuevas. En el archivo real: Platilla contabilidad las tiene en la
# fila 3, Web Adi en la fila 5 (ambas ya con los valores estáticos puestos).
PC_FILA_PLANTILLA = 3
WA_FILA_PLANTILLA = 5

# Columnas de "Platilla contabilidad" que traen una fórmula que hay que
# desplazar (arrastrar) a cada fila nueva, tal cual come en Excel.
PC_COLS_FORMULA = ["G", "I", "K", "Q", "T", "X", "Y", "Z", "AC", "AD", "AF", "AG", "AH", "AI", "AJ"]
# L es un caso especial: se desplaza como fórmula, pero solo en la primera
# fila de cada grupo de factura (las demás quedan en blanco, ver paso 22).
PC_COL_L = "L"
# Columnas "Información estática no cambia": se copia el VALOR (no fórmula)
# de la fila plantilla a cada fila nueva.
PC_COLS_ESTATICOS = ["R", "S", "U", "W", "AA", "AB"]

# Mapeo Traza -> Platilla contabilidad (letra origen -> letra destino),
# copiado como VALOR literal.
TRAZA_A_PC = {
    "AZ": "A",   # FECHA DE DEBITO -> FECHA
    "H": "B",    # SERVICIO -> CONCEPTO
    "O": "C",    # CUENTA CONTRATO -> CODIGO CLIENTE
    "A": "D",    # CECO R11 -> CCOSTO
    "G": "E",    # NOMBRE SUCURSAL -> LOCAL
    "M": "F",    # NIT -> NIT
    "T": "H",    # METODO DE PAGO -> BANCO
    "AL": "M",   # FECHA DE EMISION -> FECHA FACTURA
    "AM": "N",   # NUMERO DE FACTURA -> FACTURA
    "AQ": "O",   # VALOR ANTES DE IVA -> GASTO A CONTABILIZAR
    "AR": "P",   # IMPUESTO -> IVA
    "C": "V",    # USO FUTURO -> USO FUTURO
    "E": "AE",   # FLEX (de Traza) -> VALIDACION FLEX
}


_CELL_REF_RE = re.compile(r'(\$?)([A-Za-z]{1,3})(\$?)(\d+)')


def _desplazar_formula(formula, fila_origen, fila_destino):
    """Ajusta una fórmula de Excel como si se hubiera arrastrado (fill-down)
    de `fila_origen` a `fila_destino`: las referencias de fila relativas
    (ej. H3, N3) suben o bajan; las de fila absoluta (ej. H$3) y los rangos
    de columna completa (ej. Bancos!B:B) no se tocan."""
    if not formula or not isinstance(formula, str) or not formula.startswith("="):
        return formula
    delta = fila_destino - fila_origen
    if delta == 0:
        return formula

    def repl(m):
        dolar_col, col, dolar_fila, fila = m.groups()
        if dolar_fila == "$":
            return m.group(0)
        return f"{dolar_col}{col}{dolar_fila}{int(fila) + delta}"

    return _CELL_REF_RE.sub(repl, formula)


def _es_valor_pendiente(valor):
    """True si esta celda todavía no tiene un resultado real de Excel:
    vacía, o un error típico de fórmula (#N/A, #REF!, etc.)."""
    if valor is None:
        return True
    if isinstance(valor, str) and valor.strip().startswith("#"):
        return True
    return False


def _resolver_banco(texto, bancos):
    """Encuentra el banco de la hoja Bancos que corresponde a `texto` (el
    'método de pago' de la Traza, ej. 'DAVIVIENDA', 'BANCO DE BOGOTA
    MANUAL'), aunque no coincida letra por letra con el nombre completo
    (ej. 'BANCO DAVIVIENDA S. A.'). Devuelve (info, nombre_usado) o
    (None, None) si no encuentra ninguno o encuentra más de uno posible."""
    if not texto:
        return None, None
    clave = normalize(texto)
    if clave in bancos:
        return bancos[clave], clave

    limpio = clave
    for sufijo in (" MANUAL", " - REVISAR", " REVISAR"):
        limpio = limpio.replace(sufijo, "")
    limpio = limpio.strip()
    if limpio in bancos:
        return bancos[limpio], limpio

    candidatos = [(k, v) for k, v in bancos.items() if limpio and (limpio in k or k in limpio)]
    if len(candidatos) == 1:
        return candidatos[0][1], candidatos[0][0]
    return None, None


def _capturar_estilos(ws, fila):
    """Guarda el formato completo (relleno/color, fuente, borde, alineación,
    formato de número) de cada columna 1-20 de esa fila, para poder
    aplicárselo luego a otras filas."""
    from copy import copy
    estilos = {}
    for c in range(1, 21):
        celda = ws.cell(row=fila, column=c)
        estilos[c] = {
            "font": copy(celda.font),
            "fill": copy(celda.fill),
            "border": copy(celda.border),
            "alignment": copy(celda.alignment),
            "number_format": celda.number_format,
        }
    return estilos


def _aplicar_estilos(ws, fila, estilos):
    for c, est in estilos.items():
        celda = ws.cell(row=fila, column=c)
        celda.font = est["font"]
        celda.fill = est["fill"]
        celda.border = est["border"]
        celda.alignment = est["alignment"]
        celda.number_format = est["number_format"]


def _capturar_estilos_rango(ws, fila, col_desde, col_hasta):
    """Igual que _capturar_estilos pero para un rango de columnas cualquiera."""
    from copy import copy
    estilos = {}
    for c in range(col_desde, col_hasta + 1):
        celda = ws.cell(row=fila, column=c)
        estilos[c] = {
            "font": copy(celda.font),
            "fill": copy(celda.fill),
            "border": copy(celda.border),
            "alignment": copy(celda.alignment),
            "number_format": celda.number_format,
        }
    return estilos


def _aplicar_estilos_rango(ws, fila, estilos):
    for c, est in estilos.items():
        celda = ws.cell(row=fila, column=c)
        celda.font = est["font"]
        celda.fill = est["fill"]
        celda.border = est["border"]
        celda.alignment = est["alignment"]
        celda.number_format = est["number_format"]



def revisar_pendientes(archivo_path, hoja_traza, log, hoy=None):
    """Pantalla previa (no escribe nada): cuántas filas de Traza cumplen
    columna BA = hoy, y cuáles de esas no tienen columna AZ con una fecha
    válida."""
    hoy = hoy or date.today()
    wb = openpyxl.load_workbook(archivo_path, data_only=True)
    try:
        ws = wb[hoja_traza]
        col_az = column_index_from_string("AZ")
        col_ba = column_index_from_string("BA")
        col_am = column_index_from_string("AM")

        hoja_pc = _nombre_hoja_pc(wb)
        facturas_ya = set()
        if hoja_pc is not None:
            ws_pc = wb[hoja_pc]
            col_n = column_index_from_string("N")
            for r in range(PC_FILA_PLANTILLA, (ws_pc.max_row or 1) + 1):
                v = ws_pc.cell(row=r, column=col_n).value
                if v not in (None, ""):
                    facturas_ya.add(normalize(v))

        listas = 0
        con_error = []
        for r in range(2, (ws.max_row or 1) + 1):
            factura = normalize(ws.cell(row=r, column=col_am).value)
            if not factura or factura in facturas_ya:
                continue
            fecha_legalizacion = to_date(ws.cell(row=r, column=col_ba).value)
            if fecha_legalizacion != hoy:
                continue
            fecha_debito = ws.cell(row=r, column=col_az).value
            if not isinstance(fecha_debito, (datetime, date)):
                con_error.append(factura)
            else:
                listas += 1
    finally:
        wb.close()

    log(f"Legalizadas hoy ({hoy.isoformat()}): {listas + len(set(con_error))}.")
    return {"hoy": hoy.isoformat(), "listas": listas, "con_error": sorted(set(con_error))}


def copiar_traza_a_plantilla(archivo_path, hoja_traza, log, hoy=None):
    """Paso a paso indicado:
    1) Traza columna BA = hoy. 2) Traza columna AZ debe tener fecha (si no,
    no se copia nada). 3) Copia de Traza a Platilla contabilidad las
    columnas indicadas como VALOR, y en las columnas con fórmula la
    DESPLAZA (arrastra) a la fila nueva en vez de calcularla en Python.
    L queda en blanco en las filas duplicadas de una misma factura, y J se
    arma con la fecha (sin '/', día con cero) + el importe de esa factura."""
    hoy = hoy or date.today()
    log(f"Abriendo el archivo. Buscando filas de Traza con BA = hoy ({hoy.isoformat()})...")
    wb = openpyxl.load_workbook(archivo_path)
    ws_traza = wb[hoja_traza]

    hoja_pc = _nombre_hoja_pc(wb)
    if hoja_pc is None:
        raise ValueError(
            "El archivo no tiene la hoja de Platilla contabilidad "
            f"(se buscó como: {', '.join(PC_SHEET_ALIASES)})."
        )
    ws_pc = wb[hoja_pc]

    col_n = column_index_from_string("N")
    facturas_ya = set()
    for r in range(PC_FILA_PLANTILLA, (ws_pc.max_row or 1) + 1):
        v = ws_pc.cell(row=r, column=col_n).value
        if v not in (None, ""):
            facturas_ya.add(normalize(v))

    col_az, col_ba = column_index_from_string("AZ"), column_index_from_string("BA")
    col_am = column_index_from_string("AM")
    cols_origen = {letra: column_index_from_string(letra) for letra in TRAZA_A_PC}

    candidatas = []  # cada una: {"factura":.., "fecha_debito":.., valores por letra origen}
    for r in range(2, (ws_traza.max_row or 1) + 1):
        factura = normalize(ws_traza.cell(row=r, column=col_am).value)
        if not factura or factura in facturas_ya:
            continue
        fecha_legalizacion = to_date(ws_traza.cell(row=r, column=col_ba).value)
        if fecha_legalizacion != hoy:
            continue
        fecha_debito = ws_traza.cell(row=r, column=col_az).value
        valores = {letra: ws_traza.cell(row=r, column=c).value for letra, c in cols_origen.items()}
        candidatas.append({"factura": factura, "fecha_debito": fecha_debito, "valores": valores})

    if not candidatas:
        wb.close()
        log("No hay filas con columna BA igual a hoy (o ya están todas en Platilla contabilidad).")
        return {"rows_written": 0, "empty": True}

    # --- Paso 2: AZ debe ser fecha válida en TODAS ---
    sin_fecha = [c["factura"] for c in candidatas if not isinstance(c["fecha_debito"], (datetime, date))]
    if sin_fecha:
        log(f"ERROR: {len(sin_fecha)} factura(s) con BA = hoy no tienen columna AZ con fecha válida.")
        for f in sin_fecha[:30]:
            log(f"ERROR: Factura {f}: columna AZ (FECHA DE DEBITO) vacía o no es una fecha.")
        wb.close()
        raise ValueError(
            f"No se copió nada porque {len(sin_fecha)} factura(s) no tienen columna AZ válida. "
            "Corrige la Traza y vuelve a intentar."
        )

    # --- Agrupar por factura para el importe (L) y el número de trx (J) ---
    col_aq = cols_origen["AQ"]
    grupos = {}
    orden_grupos = []
    for c in candidatas:
        g = grupos.setdefault(c["factura"], {"suma": Decimal("0"), "fecha": c["fecha_debito"]})
        g["suma"] += _to_decimal(c["valores"]["AQ"])
        if c["factura"] not in orden_grupos:
            orden_grupos.append(c["factura"])

    j_por_factura = {}
    for factura, g in grupos.items():
        fecha_txt = g["fecha"].strftime("%d%m%Y")
        j_por_factura[factura] = fecha_txt + str(int(g["suma"]))

    # --- Paso 3: escribir en Platilla contabilidad ---
    fila_inicio = PC_FILA_PLANTILLA
    for r in range(PC_FILA_PLANTILLA, (ws_pc.max_row or 1) + 1):
        if ws_pc.cell(row=r, column=col_n).value not in (None, ""):
            fila_inicio = r + 1

    estilos_pc = _capturar_estilos_rango(ws_pc, PC_FILA_PLANTILLA, 1, 36)

    facturas_ya_escritas_en_este_lote = set()
    for i, c in enumerate(candidatas):
        fila = fila_inicio + i
        for letra_origen, letra_destino in TRAZA_A_PC.items():
            col_dest = column_index_from_string(letra_destino)
            ws_pc.cell(row=fila, column=col_dest, value=c["valores"][letra_origen])

        for letra in PC_COLS_FORMULA:
            col = column_index_from_string(letra)
            formula_plantilla = ws_pc.cell(row=PC_FILA_PLANTILLA, column=col).value
            ws_pc.cell(row=fila, column=col,
                       value=_desplazar_formula(formula_plantilla, PC_FILA_PLANTILLA, fila))

        for letra in PC_COLS_ESTATICOS:
            col = column_index_from_string(letra)
            ws_pc.cell(row=fila, column=col, value=ws_pc.cell(row=PC_FILA_PLANTILLA, column=col).value)

        # L: fórmula solo en la primera fila de cada factura del lote; en blanco después.
        col_l = column_index_from_string(PC_COL_L)
        if c["factura"] not in facturas_ya_escritas_en_este_lote:
            formula_l = ws_pc.cell(row=PC_FILA_PLANTILLA, column=col_l).value
            ws_pc.cell(row=fila, column=col_l,
                       value=_desplazar_formula(formula_l, PC_FILA_PLANTILLA, fila))
            facturas_ya_escritas_en_este_lote.add(c["factura"])
        else:
            ws_pc.cell(row=fila, column=col_l, value=None)

        # J: fecha sin "/" (día y mes con cero) + importe total de la factura.
        col_j = column_index_from_string("J")
        ws_pc.cell(row=fila, column=col_j, value=j_por_factura[c["factura"]])
        _aplicar_estilos_rango(ws_pc, fila, estilos_pc)

    log(f"'{hoja_pc}': {len(candidatas)} fila(s) copiadas desde la fila {fila_inicio}.")
    log("Recuerda abrir el archivo en Excel y guardarlo una vez, para que las fórmulas "
        "(Proveedor, Cuenta bancaria, FLEX, etc.) calculen sus valores antes de pasar a Web Adi.")

    log("Guardando el archivo...")
    try:
        guardar_atomico(wb, archivo_path)
    except PermissionError:
        raise PermissionError(
            "El archivo está abierto o en uso. Ciérralo e intenta de nuevo. No se modificó."
        )

    return {"rows_written": len(candidatas), "empty": False}


# ---------------------------------------------------------------------------
# Platilla contabilidad -> Web Adi
# ---------------------------------------------------------------------------

# Débitos: letra en Platilla contabilidad -> letra en Web Adi (valor literal).
PC_A_WA_DEBITO = {
    "W": "B",   # CTA CONTABLE -> Cuenta
    "X": "C",   # SUC -> Sucursal
    "Y": "D",   # CC R12 -> Centro Costos
    "Z": "E",   # PV -> Punto Venta
    "V": "J",   # USO FUTURO -> Uso Futuro
    "O": "K",   # GASTO A CONTABILIZAR -> Débito
    "Q": "M",   # OBSERVACION ORACLE -> Descripción de Línea
    "G": "R",   # PROVEEDOR -> Info de Línea Adicional 3 (nombre del tercero)
    "A": "T",   # FECHA -> FECHA
}
# Estáticos de Web Adi (se toman de su propia fila plantilla, fila 5).
WA_COLS_ESTATICOS = ["A", "F", "G", "H", "I", "O", "S"]

# Formato fijo SOLO para las líneas de crédito (las que empiezan con cuenta
# "11..."): estos valores no salen de ninguna fórmula ni de la Traza, son
# siempre los mismos en toda línea de crédito.
WA_CREDITO_ESTATICOS = {"C": "000", "D": "00000", "E": "00000", "J": "0000"}

# Créditos: letra en Platilla contabilidad -> letra en Web Adi.
PC_A_WA_CREDITO = {
    "AH": "B",  # CUENTA CONTABLE (del banco) -> Cuenta
    "L": "L",   # IMPORTE -> Crédito
    "Q": ("M", "N"),  # OBSERVACION ORACLE -> Descripción de Línea Y también Informe Detallado
    "AI": "P",  # NIT BANCO -> TRAER (NIT sin dígito)
    "H": "R",   # BANCO -> nombre del tercero
    "A": "T",   # FECHA -> FECHA
}


def _fila_ultima_con_valor(ws, col_letra, desde_fila):
    col = column_index_from_string(col_letra)
    ultima = desde_fila - 1
    for r in range(desde_fila, (ws.max_row or 1) + 1):
        if ws.cell(row=r, column=col).value not in (None, ""):
            ultima = r
    return ultima


def copiar_plantilla_a_webadi(archivo_path, log):
    """Reconstruye por completo el bloque de datos de 'Web Adi' (débitos y
    luego créditos) a partir de lo que HOY hay calculado en 'Platilla
    contabilidad'. Requiere que el archivo ya se haya abierto y guardado en
    Excel después de 'Copiar a Platilla contabilidad', para que las
    fórmulas (Proveedor, Cuenta bancaria, FLEX, etc.) ya tengan un valor
    calculado."""
    log("Abriendo el archivo (con los valores ya calculados por Excel)...")
    wb_valores = openpyxl.load_workbook(archivo_path, data_only=True)
    try:
        hoja_pc = _nombre_hoja_pc(wb_valores)
        if hoja_pc is None:
            raise ValueError(
                "El archivo no tiene la hoja de Platilla contabilidad "
                f"(se buscó como: {', '.join(PC_SHEET_ALIASES)})."
            )
        if "Web Adi" not in wb_valores.sheetnames:
            raise ValueError("El archivo no tiene la hoja 'Web Adi'.")
        if "Bancos" not in wb_valores.sheetnames:
            raise ValueError("El archivo no tiene la hoja 'Bancos'.")

        ws_pc = wb_valores[hoja_pc]
        ws_bancos = wb_valores["Bancos"]

        # Maestro de bancos completo (nombre -> cuenta contable, NIT sin
        # dígito y dígito), para resolver el banco de cada crédito en
        # Python directamente, sin depender de que el texto de "método de
        # pago" en la Traza coincida letra por letra con el nombre en Bancos.
        bancos = {}
        for row in ws_bancos.iter_rows(min_row=2, max_row=ws_bancos.max_row, min_col=1, max_col=5, values_only=True):
            a, b, c, d, e = row
            if b is None:
                continue
            nit_dash = "" if e is None else str(e).strip()
            digito = nit_dash.split("-")[-1].strip() if "-" in nit_dash else ""
            bancos[normalize(b)] = {
                "cuenta_contable": "" if d is None else str(d).strip(),
                "nit_sin_digito": "" if a is None else str(a).strip(),
                "digito": digito,
                "nombre": str(b).strip(),
            }

        col_n = column_index_from_string("N")
        ultima_pc = _fila_ultima_con_valor(ws_pc, "N", PC_FILA_PLANTILLA)
        if ultima_pc < PC_FILA_PLANTILLA:
            log("No hay filas con datos en Platilla contabilidad.")
            return {"debitos": 0, "creditos": 0, "empty": True}

        filas_pc = []
        faltantes = []
        for r in range(PC_FILA_PLANTILLA, ultima_pc + 1):
            if ws_pc.cell(row=r, column=col_n).value in (None, ""):
                continue
            valores = {letra: ws_pc.cell(row=r, column=column_index_from_string(letra)).value
                       for letra in ("A", "F", "G", "H", "J", "L", "O", "Q", "V", "W", "X", "Y", "Z", "AH", "AI")}
            # G/X/Y/Z son fórmulas (Proveedor y Centros de Costos) que de
            # verdad necesitan que Excel ya haya recalculado; si Excel no
            # encontró coincidencia, el valor queda como texto de error
            # (#N/A, #REF!...) en vez de None, así que se revisan ambos casos.
            if any(_es_valor_pendiente(valores[c]) for c in ("G", "X", "Y", "Z")):
                faltantes.append(str(ws_pc.cell(row=r, column=col_n).value))
            filas_pc.append(valores)

        if faltantes:
            raise ValueError(
                f"{len(faltantes)} fila(s) de Platilla contabilidad todavía no tienen las fórmulas "
                "calculadas (Proveedor, Centros de Costos...). Abre el archivo en Excel, "
                "espera a que recalcule, guárdalo, y vuelve a intentar. Facturas: "
                + ", ".join(faltantes[:15])
            )

        # --- Filas de crédito: solo donde quedó IMPORTE (L) ---
        filas_credito_pc = [v for v in filas_pc if v["L"] not in (None, "")]

        # Resolver el banco de cada crédito en Python (no depende de que
        # Excel haya encontrado la coincidencia exacta del nombre).
        bancos_sin_resolver = set()
        for v in filas_credito_pc:
            info, _ = _resolver_banco(v["H"], bancos)
            v["_banco_info"] = info
            if info is None:
                bancos_sin_resolver.add(str(v["H"]))

        if bancos_sin_resolver:
            log(f"ALERTA: {len(bancos_sin_resolver)} nombre(s) de banco/método de pago no coinciden "
                f"con ningún banco de la hoja Bancos: {sorted(bancos_sin_resolver)}. "
                "Esas filas no generarán línea de crédito hasta que agregues ese nombre "
                "(o uno parecido) en la hoja Bancos.")
            filas_credito_pc = [v for v in filas_credito_pc if v["_banco_info"] is not None]
    finally:
        wb_valores.close()

    log(f"Débitos a escribir: {len(filas_pc)} · Créditos a escribir: {len(filas_credito_pc)}.")

    # --- Reescribir Web Adi conservando fórmulas/estilos, con un libro normal ---
    wb = openpyxl.load_workbook(archivo_path, data_only=False)
    ws_wa = wb["Web Adi"]

    ultima_actual = WA_FILA_PLANTILLA - 1
    for r in range(WA_FILA_PLANTILLA, (ws_wa.max_row or 1) + 1):
        if any(ws_wa.cell(row=r, column=c).value not in (None, "") for c in range(1, 21)):
            ultima_actual = r

    # Antes de borrar nada: capturar el color/formato de una fila de débito y
    # una de crédito ya existentes, para poder aplicárselo a las filas nuevas
    # (si cambia cuántas filas hay de cada tipo, el color debe seguir siendo
    # el correcto de cada sección, no el que quedó pegado en esa posición).
    fila_credito_existente = None
    col_credito = column_index_from_string("L")
    for r in range(WA_FILA_PLANTILLA, ultima_actual + 1):
        if ws_wa.cell(row=r, column=col_credito).value not in (None, ""):
            fila_credito_existente = r
            break

    estilos_debito = _capturar_estilos(ws_wa, WA_FILA_PLANTILLA)
    estilos_credito = _capturar_estilos(ws_wa, fila_credito_existente) if fila_credito_existente else estilos_debito
    if fila_credito_existente is None:
        log("ALERTA: no encontré ninguna fila de crédito existente para copiar su color; "
            "las filas de crédito quedaron con el mismo color que las de débito.")

    # Limpiar todo el bloque de datos antes de reescribir.
    for r in range(WA_FILA_PLANTILLA, max(ultima_actual, WA_FILA_PLANTILLA) + 1):
        for c in range(1, 21):
            ws_wa.cell(row=r, column=c, value=None)

    estaticos_wa = {letra: ws_wa.cell(row=WA_FILA_PLANTILLA, column=column_index_from_string(letra)).value
                    for letra in WA_COLS_ESTATICOS}
    formula_n = ws_wa.cell(row=WA_FILA_PLANTILLA, column=column_index_from_string("N")).value

    fila = WA_FILA_PLANTILLA
    # Débitos
    for v in filas_pc:
        for letra, valor in estaticos_wa.items():
            ws_wa.cell(row=fila, column=column_index_from_string(letra), value=valor)
        for letra_o, letra_d in PC_A_WA_DEBITO.items():
            ws_wa.cell(row=fila, column=column_index_from_string(letra_d), value=v[letra_o])
        nit_completo = str(v["F"]) if v["F"] is not None else ""
        nit_sin_digito, _, digito = nit_completo.partition("-")
        ws_wa.cell(row=fila, column=column_index_from_string("P"), value=nit_sin_digito)
        ws_wa.cell(row=fila, column=column_index_from_string("Q"), value=digito)
        ws_wa.cell(row=fila, column=column_index_from_string("N"),
                   value=_desplazar_formula(formula_n, WA_FILA_PLANTILLA, fila))
        _aplicar_estilos(ws_wa, fila, estilos_debito)
        fila += 1

    # Créditos
    for v in filas_credito_pc:
        for letra, valor in estaticos_wa.items():
            ws_wa.cell(row=fila, column=column_index_from_string(letra), value=valor)
        for letra, valor in WA_CREDITO_ESTATICOS.items():
            ws_wa.cell(row=fila, column=column_index_from_string(letra), value=valor)
        banco_nombre = v["H"]
        banco_info = v["_banco_info"]
        ws_wa.cell(row=fila, column=column_index_from_string("B"), value=banco_info["cuenta_contable"])
        ws_wa.cell(row=fila, column=column_index_from_string("L"), value=v["L"])
        ws_wa.cell(row=fila, column=column_index_from_string("M"), value=v["Q"])
        ws_wa.cell(row=fila, column=column_index_from_string("N"), value=v["Q"])
        ws_wa.cell(row=fila, column=column_index_from_string("P"), value=banco_info["nit_sin_digito"])
        ws_wa.cell(row=fila, column=column_index_from_string("Q"), value=banco_info["digito"])
        ws_wa.cell(row=fila, column=column_index_from_string("R"), value=banco_nombre)
        ws_wa.cell(row=fila, column=column_index_from_string("T"), value=v["A"])
        _aplicar_estilos(ws_wa, fila, estilos_credito)
        fila += 1

    log("Guardando el archivo...")
    try:
        guardar_atomico(wb, archivo_path)
    except PermissionError:
        raise PermissionError(
            "El archivo está abierto o en uso. Ciérralo e intenta de nuevo. No se modificó."
        )

    return {"debitos": len(filas_pc), "creditos": len(filas_credito_pc), "empty": False}


# Puente expuesto a la interfaz JS (pywebview)
# ---------------------------------------------------------------------------

class Puente:
    def __init__(self):
        self.window = None

    def _eval(self, js):
        if self.window is None:
            return
        try:
            self.window.evaluate_js(js)
        except Exception:
            print("(ventana no disponible para actualizar)")

    def _log(self, msg):
        print("[Integración]", msg)
        self._eval(f"window.appendLog({json.dumps(msg)})")

    def _log_cxp(self, msg):
        print("[CXP]", msg)
        self._eval(f"window.appendLogCXP({json.dumps(msg)})")

    def _log_webadi(self, msg):
        print("[WebADI]", msg)
        self._eval(f"window.appendLogWebAdi({json.dumps(msg)})")

    # -- Estado inicial módulo 1 --

    def get_initial_state(self):
        cfg = load_config()
        state = {
            "traza_path": cfg.get("traza_path", ""),
            "azimut_path": cfg.get("azimut_path", ""),
            "hoja_destino": cfg.get("hoja_destino", ""),
            "sheets": [],
        }
        if state["traza_path"] and os.path.exists(state["traza_path"]):
            try:
                wb = openpyxl.load_workbook(state["traza_path"], read_only=True)
                state["sheets"] = wb.sheetnames
                wb.close()
            except Exception:
                state["traza_path"] = ""
        return state

    # -- Selectores de archivo (reutilizados por ambos módulos) --

    def pick_traza(self):
        try:
            result = self.window.create_file_dialog(
                webview.OPEN_DIALOG,
                file_types=("Archivos Excel (*.xlsx;*.xlsm)", "Todos los archivos (*.*)"),
            )
            if not result:
                return {"ok": False}
            path = result[0]
            wb = openpyxl.load_workbook(path, read_only=True)
            sheets = wb.sheetnames
            wb.close()
            return {"ok": True, "path": path, "sheets": sheets}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def pick_azimut(self):
        try:
            result = self.window.create_file_dialog(
                webview.OPEN_DIALOG,
                file_types=("Archivos Excel (*.xlsx;*.xlsm)", "Todos los archivos (*.*)"),
            )
            if not result:
                return {"ok": False}
            return {"ok": True, "path": result[0]}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # -- Módulo 1: ejecutar --

    def ejecutar(self, traza_path, hoja, azimut_path):
        def worker():
            try:
                resumen = copy_azimut_to_traza(traza_path, hoja, azimut_path, self._log)
                cfg = load_config()
                cfg.update({"traza_path": traza_path, "azimut_path": azimut_path, "hoja_destino": hoja})
                save_config(cfg)
                self._log("")
                self._log("Proceso completado correctamente.")
                payload = {"ok": True}
                payload.update(resumen)
                self._eval(f"window.procesoTerminado({json.dumps(payload)})")
            except PermissionError as e:
                self._log(f"ERROR: {e}")
                self._eval(f"window.procesoTerminado({json.dumps({'ok': False, 'error': str(e)})})")
            except Exception as e:
                traceback.print_exc()
                self._log(f"ERROR inesperado: {e}")
                self._eval(f"window.procesoTerminado({json.dumps({'ok': False, 'error': str(e)})})")
        threading.Thread(target=worker, daemon=True).start()
        return {"started": True}

    # -- Estado inicial módulo 2 --

    def get_initial_state_cxp(self):
        cfg = load_config()
        state = {
            "traza_path": cfg.get("cxp_traza_path", ""),
            "hoja": cfg.get("cxp_hoja", ""),
            "sheets": [],
        }
        if state["traza_path"] and os.path.exists(state["traza_path"]):
            try:
                wb = openpyxl.load_workbook(state["traza_path"], read_only=True)
                state["sheets"] = wb.sheetnames
                wb.close()
            except Exception:
                state["traza_path"] = ""
        return state

    # -- Módulo 2: generar CXP --

    def get_metodos_pago(self):
        """Devuelve la lista fija de métodos de pago disponibles."""
        return {"ok": True, "metodos": CXP_METODOS_DISPONIBLES}

    def generar_cxp(self, traza_path, hoja, metodos_filtro):
        sugerido = f"CXP_{datetime.now().strftime('%d%m%Y')}.xlsx"
        cfg = load_config()

        # El diálogo de guardado (que bloquea el hilo principal) se abre SOLO
        # si hay registros que exportar. Por eso primero hacemos el filtrado
        # en background, y desde allí pedimos la ruta si hace falta.
        def worker():
            try:
                # Paso 1: analizar una sola vez. El resultado se reutiliza al generar.
                pre = generar_plantilla_cxp(traza_path, hoja, None, metodos_filtro, self._log_cxp)
                preparado = pre.pop("_preparado", None)
                if pre.get("empty"):
                    self._eval(f"window.procesoTerminadoCXP({json.dumps({'ok': True, **pre})})")
                    return

                # Paso 2: hay registros → pedir ruta (debe correr en hilo principal)
                ruta_holder = []

                def pedir_ruta():
                    try:
                        res = self.window.create_file_dialog(
                            webview.SAVE_DIALOG,
                            directory=cfg.get("cxp_save_dir", ""),
                            save_filename=sugerido,
                            file_types=("Archivos Excel (*.xlsx)",),
                        )
                        ruta_holder.append(res)
                    except Exception as e:
                        ruta_holder.append(None)

                import threading as _th
                ev = _th.Event()
                def pedir_y_avisar():
                    pedir_ruta(); ev.set()

                self.window.evaluate_js("null")  # warm-up por si acaso
                t = _th.Thread(target=pedir_y_avisar, daemon=True)
                t.start(); ev.wait(timeout=120)

                res = ruta_holder[0] if ruta_holder else None
                if not res:
                    self._eval(f"window.procesoTerminadoCXP({json.dumps({'ok': False, 'cancelled': True})})")
                    return

                salida_path = res[0] if isinstance(res, (list, tuple)) else res
                if not salida_path.lower().endswith(".xlsx"):
                    salida_path += ".xlsx"

                # Paso 3: generar con los datos ya analizados; no se vuelve a recorrer la Traza.
                resumen = generar_plantilla_cxp(
                    traza_path, hoja, salida_path, metodos_filtro, self._log_cxp,
                    preparado=preparado,
                )
                nueva_cfg = load_config()
                nueva_cfg.update({
                    "cxp_traza_path": traza_path,
                    "cxp_hoja": hoja,
                    "cxp_save_dir": os.path.dirname(salida_path),
                })
                save_config(nueva_cfg)
                payload = {"ok": True}
                payload.update(resumen)
                self._eval(f"window.procesoTerminadoCXP({json.dumps(payload)})")

            except PermissionError as e:
                self._log_cxp(f"ERROR: {e}")
                self._eval(f"window.procesoTerminadoCXP({json.dumps({'ok': False, 'error': str(e)})})")
            except Exception as e:
                traceback.print_exc()
                self._log_cxp(f"ERROR inesperado: {e}")
                self._eval(f"window.procesoTerminadoCXP({json.dumps({'ok': False, 'error': str(e)})})")

        threading.Thread(target=worker, daemon=True).start()
        return {"ok": True, "started": True}

    # -- Módulo 3: Contabilización (funciona como una macro: --
    # -- lee y escribe sobre el MISMO archivo, sin diálogo de guardado) --

    def get_initial_state_webadi(self):
        cfg = load_config()
        state = {
            "archivo_path": cfg.get("webadi_archivo_path", ""),
            "hoja": cfg.get("webadi_hoja", ""),
            "sheets": [],
        }
        if state["archivo_path"] and os.path.exists(state["archivo_path"]):
            try:
                wb = openpyxl.load_workbook(state["archivo_path"], read_only=True)
                state["sheets"] = wb.sheetnames
                wb.close()
            except Exception:
                state["archivo_path"] = ""
        return state

    def revisar_pendientes_ui(self, archivo_path, hoja):
        try:
            resumen = revisar_pendientes(archivo_path, hoja, self._log_webadi)
            return {"ok": True, **resumen}
        except Exception as e:
            traceback.print_exc()
            return {"ok": False, "error": str(e)}

    def copiar_pendientes_ui(self, archivo_path, hoja):
        def worker():
            try:
                resumen = copiar_traza_a_plantilla(archivo_path, hoja, self._log_webadi)
                cfg = load_config()
                cfg.update({"webadi_archivo_path": archivo_path, "webadi_hoja": hoja})
                save_config(cfg)
                payload = {"ok": True}
                payload.update(resumen)
                self._eval(f"window.procesoTerminadoWebAdi({json.dumps(payload)})")
            except PermissionError as e:
                self._log_webadi(f"ERROR: {e}")
                self._eval(f"window.procesoTerminadoWebAdi({json.dumps({'ok': False, 'error': str(e)})})")
            except ValueError as e:
                self._eval(f"window.procesoTerminadoWebAdi({json.dumps({'ok': False, 'error': str(e)})})")
            except Exception as e:
                traceback.print_exc()
                self._log_webadi(f"ERROR inesperado: {e}")
                self._eval(f"window.procesoTerminadoWebAdi({json.dumps({'ok': False, 'error': str(e)})})")

        threading.Thread(target=worker, daemon=True).start()
        return {"ok": True, "started": True}

    def copiar_a_webadi_ui(self, archivo_path):
        def worker():
            try:
                resumen = copiar_plantilla_a_webadi(archivo_path, self._log_webadi)
                payload = {"ok": True}
                payload.update(resumen)
                self._eval(f"window.procesoTerminadoWebAdi2({json.dumps(payload)})")
            except PermissionError as e:
                self._log_webadi(f"ERROR: {e}")
                self._eval(f"window.procesoTerminadoWebAdi2({json.dumps({'ok': False, 'error': str(e)})})")
            except ValueError as e:
                self._eval(f"window.procesoTerminadoWebAdi2({json.dumps({'ok': False, 'error': str(e)})})")
            except Exception as e:
                traceback.print_exc()
                self._log_webadi(f"ERROR inesperado: {e}")
                self._eval(f"window.procesoTerminadoWebAdi2({json.dumps({'ok': False, 'error': str(e)})})")

        threading.Thread(target=worker, daemon=True).start()
        return {"ok": True, "started": True}


def main():
    puente = Puente()
    window = webview.create_window(
        APP_TITLE,
        gui_path("index.html"),
        js_api=puente,
        width=1000,
        height=860,
        min_size=(820, 700),
        background_color="#F4F5F7",
    )
    puente.window = window
    webview.start()


if __name__ == "__main__":
    main()

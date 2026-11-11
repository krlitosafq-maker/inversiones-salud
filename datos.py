"""Descarga los archivos de Datos Abiertos del MEF y guarda las inversiones
de la funcion elegida (Salud) en una base de datos local SQLite."""
import csv
import datetime
import io
import json
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile

CONGELADO = getattr(sys, "frozen", False)  # True cuando corre como .exe
# Carpeta donde se guardan los datos: junto al .exe o junto a este archivo
BASE = os.path.dirname(os.path.abspath(sys.executable if CONGELADO else __file__))
# Carpeta de los archivos internos (pantalla)
RECURSOS = getattr(sys, "_MEIPASS", BASE)
CARPETA_DATOS = os.path.join(BASE, "data")
RUTA_BD = os.path.join(CARPETA_DATOS, "inversiones.db")

URL_ARCHIVOS = "https://fs.datosabiertos.mef.gob.pe/datastorefiles/"
FUNCION = "SALUD"          # funcion del Banco de Inversiones
FUNCION_CODIGO = "20"      # la misma funcion en el clasificador presupuestal
AGENTE = "Mozilla/5.0 (seguimiento-inversiones)"

NUMERICAS = re.compile(
    r"^(MONTO|DEV|PIA|PIM|CERTIF|COMPROM|SALDO|AVANCE|PMI_ANIO|COSTO|CTRL|LATITUD|LONGITUD|NUM_)"
)

csv.field_size_limit(100_000_000)


def conectar():
    os.makedirs(CARPETA_DATOS, exist_ok=True)
    con = sqlite3.connect(RUTA_BD, timeout=60)
    con.row_factory = sqlite3.Row
    con.execute("CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)")
    return con


def hay_datos(con):
    return con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='inversiones'"
    ).fetchone() is not None


def _columna(nombre):
    n = nombre.strip().upper().replace("Ñ", "N")
    return re.sub(r"[^A-Z0-9_]", "_", n)


def _numero(texto):
    if texto is None or texto == "":
        return None
    try:
        return float(texto)
    except ValueError:
        return None


class _Conteo(io.RawIOBase):
    """Envuelve la descarga para saber cuanto se ha leido."""

    def __init__(self, resp, avisar):
        self.resp = resp
        self.avisar = avisar
        self.leido = 0
        self.total = int(resp.headers.get("Content-Length") or 0)
        self._ultimo = 0

    def readable(self):
        return True

    def readinto(self, b):
        datos = self.resp.read(len(b))
        n = len(datos)
        b[:n] = datos
        self.leido += n
        if self.leido - self._ultimo > 4_000_000 or n == 0:
            self._ultimo = self.leido
            self.avisar(self.leido, self.total)
        return n


def _abrir(url):
    pedido = urllib.request.Request(url, headers={"User-Agent": AGENTE})
    return urllib.request.urlopen(pedido, timeout=120)


HOST_INVIERTE = "ofi5.mef.gob.pe"
URL_IDEAS = "https://%s/inviertePub/ConsultaPublica/exportarConsultaAvanzadaIdeas" % HOST_INVIERTE
# Mismos parametros que envia la pagina publica "Consulta avanzada de ideas" al buscar por funcion
FILTRO_IDEAS = {
    "filters": "", "ip": "", "cboNom": "1", "txtNom": "", "txtCod_unico": "0", "txtCod_idea": "0",
    "cboPais": "0", "cboDpto": "0", "cboProv": "0", "cboDist": "0", "optUf": "*", "cboGNSect": "*",
    "cboGNPlie": "", "cboGNUF": "", "cboGR": "*", "cboGRUf": "", "optGL": "", "cboGLDpto": "*",
    "cboGLProv": "*", "cboGLDist": "*", "cboGLUf": "", "cboGLManPlie": "*", "cboGLManUf": "",
    "cboSitu": "0", "cboNivReqViab": "*", "cboEstu": "*", "cboEsta": "*", "optFecha": "*",
    "txtIni": "", "txtFin": "", "chkMonto": False, "txtMin": "", "txtMax": "",
    "chkMontoElaboracion": False, "txtMinME": "", "txtMaxME": "", "tipo": "1",
    "cboFunc": FUNCION_CODIGO, "chkInactivo": "0", "cboDivision": "0", "cboGrupo": "0",
    "rbtnCadena": "", "isSearch": False, "sortField": "", "sortOrder": "desc", "optUG": "DEP",
    "chkFoniprel": "",
}
CAMPOS_IDEA = (  # columna local, titulo de la columna en el Excel del MEF, es numero
    ("cod_idea", "Código idea", False), ("nombre", "Nombre de la idea", False),
    ("monto_inversion", "Monto tentativo de la Inversión", True),
    ("costo_estudio", "Costo de la Elaboración técnico", True),
    ("tipo", "Tipo de Inversión", False), ("nivel", "Nivel de Gobierno", False),
    ("sector", "Sector", False), ("entidad", "Entidad", False),
    ("funcion", "Función", False), ("division", "División", False), ("grupo", "Grupo", False),
    ("enfoque", "Enfoque", False),
    ("cod_unico", "Código Unico de Inversiones", False),
    ("fecha_registro", "Fecha de registro", False),
    ("estado", "Estado de la idea de inversión", False),
    ("ejecucion", "Descripción de Ejecución", False), ("anio", "Año de ejecución", False),
    ("departamento", "Departamento", False),
)
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _descargar_ideas(destino):
    """Pide al MEF el Excel de la consulta avanzada de ideas. Se usa curl.exe de Windows porque
    el servidor del MEF no envia su cadena completa de certificados y Windows si sabe validarla."""
    curl = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "curl.exe")
    if not os.path.exists(curl):
        raise RuntimeError("Esta version de Windows no incluye curl.exe, necesario para las ideas")
    try:
        r = subprocess.run(
            [curl, "--silent", "--show-error", "--fail", "--max-time", "300",
             "-A", AGENTE, "-H", "Content-Type: application/json; charset=utf-8",
             "--data-binary", "@-", "-o", destino, URL_IDEAS],
            input=json.dumps(FILTRO_IDEAS).encode("utf-8"), capture_output=True, timeout=320,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired:
        raise RuntimeError("La consulta de ideas del MEF no respondio a tiempo")
    if r.returncode != 0:
        raise RuntimeError("No se pudo consultar las ideas en el MEF: "
                           + r.stderr.decode("utf-8", "replace").strip()[:200])


def _filas_excel(ruta):
    """Lee la primera hoja de un .xlsx y entrega cada fila como {letra de columna: valor}."""
    with zipfile.ZipFile(ruta) as z:
        textos = []
        if "xl/sharedStrings.xml" in z.namelist():
            with z.open("xl/sharedStrings.xml") as f:
                for _, e in ET.iterparse(f):
                    if e.tag == _NS + "si":
                        textos.append("".join(x.text or "" for x in e.iter(_NS + "t")))
                        e.clear()
        with z.open("xl/worksheets/sheet1.xml") as f:
            for _, e in ET.iterparse(f):
                if e.tag != _NS + "row":
                    continue
                fila = {}
                for c in e.findall(_NS + "c"):
                    letra = re.match(r"[A-Z]+", c.get("r", "")).group(0)
                    v = c.find(_NS + "v")
                    if c.get("t") == "s" and v is not None:
                        fila[letra] = textos[int(v.text)]
                    elif c.get("t") == "inlineStr":
                        fila[letra] = "".join(x.text or "" for x in c.iter(_NS + "t"))
                    elif v is not None:
                        fila[letra] = v.text
                e.clear()
                yield fila


def _fecha_iso(texto):
    """Convierte '12/03/2018' en '2018-03-12'."""
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", texto or "")
    return "%s-%s-%s" % (m.group(3), m.group(2), m.group(1)) if m else (texto or "")


def actualizar_ideas(avisar):
    """Descarga las ideas de inversion de la funcion desde la consulta publica del Banco de Inversiones."""
    avisar("Ideas de inversion (Banco de Inversiones)", 0.1)
    temporal = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temporal.close()
    try:
        _descargar_ideas(temporal.name)
        avisar("Ideas de inversion (Banco de Inversiones)", 0.7)
        if not zipfile.is_zipfile(temporal.name):
            raise RuntimeError("La consulta de ideas del MEF no devolvio un archivo Excel")
        columnas, ideas = None, {}
        unir = [i for i, c in enumerate(CAMPOS_IDEA) if c[0] in ("departamento", "ejecucion")]
        for fila in _filas_excel(temporal.name):
            if columnas is None:
                por_titulo = {str(v).strip(): k for k, v in fila.items()}
                if CAMPOS_IDEA[0][1] in por_titulo:
                    faltan = [t for _, t, _ in CAMPOS_IDEA if t not in por_titulo]
                    if faltan:
                        raise RuntimeError("El Excel de ideas cambio de formato: faltan " + ", ".join(faltan))
                    columnas = [por_titulo[t] for _, t, _ in CAMPOS_IDEA]
                continue
            valores = []
            for letra, (campo, _, es_num) in zip(columnas, CAMPOS_IDEA):
                v = (fila.get(letra) or "").strip()
                if es_num:
                    v = _numero(v.replace(",", ""))
                elif campo == "fecha_registro":
                    v = _fecha_iso(v)
                valores.append(v)
            if not valores[0]:
                continue
            # El MEF repite la idea por cada departamento o ejecutora: se junta en un solo registro
            previa = ideas.get(valores[0])
            if previa is None:
                ideas[valores[0]] = valores
            else:
                for i in unir:
                    if valores[i] and valores[i] not in previa[i].split("; "):
                        previa[i] = (previa[i] + "; " + valores[i]) if previa[i] else valores[i]
    finally:
        try:
            os.unlink(temporal.name)
        except OSError:
            pass
    if not ideas:
        raise RuntimeError("La consulta de ideas del MEF no trajo registros")

    con = conectar()
    try:
        con.execute("BEGIN")
        con.execute("DROP TABLE IF EXISTS ideas")
        con.execute("CREATE TABLE ideas (%s)" % ", ".join(
            "%s %s" % (c, "REAL" if es_num else "TEXT") for c, _, es_num in CAMPOS_IDEA))
        con.executemany("INSERT INTO ideas VALUES (%s)" % ",".join("?" * len(CAMPOS_IDEA)),
                        list(ideas.values()))
        # En el Excel del MEF las columnas Funcion, Division y Grupo vienen corridas una posicion
        con.execute("UPDATE ideas SET grupo = division, division = funcion, funcion = grupo "
                    "WHERE grupo = ? AND funcion <> ?", (FUNCION, FUNCION))
        con.execute("CREATE INDEX ix_ideas_cui ON ideas (cod_unico)")
        con.execute("CREATE INDEX ix_ideas_cod ON ideas (cod_idea)")
        ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        for k, v in (("ideas", str(len(ideas))), ("ideas_actualizado", ahora), ("ideas_error", "")):
            con.execute("INSERT OR REPLACE INTO meta (k, v) VALUES (?, ?)", (k, v))
        con.commit()
    except BaseException:
        con.rollback()
        raise
    finally:
        con.close()
    return len(ideas)


def _cargar(con, tabla, lector, filtro, indices):
    """Crea la tabla a partir de la cabecera del CSV y guarda las filas que pasan el filtro."""
    cabecera = [_columna(c) for c in next(lector)]
    pos = {c: i for i, c in enumerate(cabecera)}
    numericas = [bool(NUMERICAS.match(c)) for c in cabecera]
    columnas = ", ".join(
        '"%s" %s' % (c.lower(), "REAL" if es_num else "TEXT")
        for c, es_num in zip(cabecera, numericas)
    )
    con.execute('DROP TABLE IF EXISTS "%s"' % tabla)
    con.execute('CREATE TABLE "%s" (%s)' % (tabla, columnas))
    insertar = 'INSERT INTO "%s" VALUES (%s)' % (tabla, ",".join("?" * len(cabecera)))
    acepta = filtro(pos)
    lote, guardadas, n = [], 0, len(cabecera)
    for fila in lector:
        if len(fila) != n or not acepta(fila):
            continue
        lote.append([
            _numero(v) if es_num else v.strip() for v, es_num in zip(fila, numericas)
        ])
        if len(lote) >= 2000:
            con.executemany(insertar, lote)
            guardadas += len(lote)
            lote = []
    if lote:
        con.executemany(insertar, lote)
        guardadas += len(lote)
    for col in indices:
        if col.upper() in pos:
            con.execute('CREATE INDEX "ix_%s_%s" ON "%s" ("%s")' % (tabla, col, tabla, col))
    return guardadas, pos


def _exigir(pos, columnas, archivo):
    faltan = [c for c in columnas if c not in pos]
    if faltan:
        raise RuntimeError(
            "El archivo %s cambio de formato: faltan las columnas %s" % (archivo, ", ".join(faltan))
        )


def actualizar(avisar):
    """Actualiza inversiones e ideas. Si las ideas fallan, las inversiones igual quedan actualizadas."""
    resultado = _actualizar_inversiones(avisar)
    try:
        resultado["ideas"] = actualizar_ideas(avisar)
    except Exception as e:
        resultado["aviso"] = "Las inversiones se actualizaron, pero no las ideas de inversion: %s" % e
        con = conectar()
        try:
            con.execute("INSERT OR REPLACE INTO meta (k, v) VALUES ('ideas_error', ?)", (str(e),))
            con.commit()
        finally:
            con.close()
    return resultado


def _actualizar_inversiones(avisar):
    """Descarga los archivos de Datos Abiertos y reemplaza los datos locales.
    Si algo falla, se conservan los anteriores."""
    con = conectar()
    try:
        con.execute("BEGIN")

        # 1. Detalle de inversiones (Banco de Inversiones)
        def progreso(nombre):
            def f(leido, total):
                avisar(nombre, leido / total if total else 0)
            return f

        archivo = "DETALLE_INVERSIONES.csv"
        with _abrir(URL_ARCHIVOS + archivo) as resp:
            fecha_detalle = resp.headers.get("Last-Modified", "")
            texto = io.TextIOWrapper(
                io.BufferedReader(_Conteo(resp, progreso("Inversiones (Banco de Inversiones)"))),
                encoding="utf-8-sig", newline="")

            def filtro_detalle(pos):
                _exigir(pos, ["FUNCION", "CODIGO_UNICO", "CODIGO_SNIP", "PIM_ANIO_ACTUAL",
                              "DEV_ANIO_ACTUAL", "NOMBRE_INVERSION"], archivo)
                i = pos["FUNCION"]
                return lambda fila: fila[i].strip() == FUNCION

            total_inv, pos = _cargar(con, "inversiones", csv.reader(texto), filtro_detalle,
                                     ["codigo_unico", "departamento", "nivel"])
        if total_inv == 0:
            raise RuntimeError("El archivo de inversiones no trajo registros de la funcion " + FUNCION)

        con.execute("ALTER TABLE inversiones ADD COLUMN clave TEXT")
        con.execute(
            "UPDATE inversiones SET clave = CASE WHEN codigo_unico <> '' THEN codigo_unico "
            "ELSE 'SNIP-' || codigo_snip END")
        con.execute("CREATE INDEX ix_inversiones_clave ON inversiones (clave)")
        cuis = {f[0] for f in con.execute("SELECT codigo_unico FROM inversiones WHERE codigo_unico <> ''")}
        anio = con.execute("SELECT MAX(anio_proceso) FROM inversiones").fetchone()[0] if "ANIO_PROCESO" in pos else None
        anio = str(anio or datetime.date.today().year)

        # 2. Formato 12-B (seguimiento mensual)
        archivo = "FORMATO_12B.csv"
        with _abrir(URL_ARCHIVOS + archivo) as resp:
            texto = io.TextIOWrapper(
                io.BufferedReader(_Conteo(resp, progreso("Formato 12-B (seguimiento mensual)"))),
                encoding="utf-8-sig", newline="")

            def filtro_f12b(pos):
                _exigir(pos, ["CODIGO_UNICO"], archivo)
                i = pos["CODIGO_UNICO"]
                return lambda fila: fila[i].strip() in cuis

            total_f12b, _ = _cargar(con, "f12b", csv.reader(texto), filtro_f12b, ["codigo_unico"])

        # 3. Ejecucion presupuestal por proyecto (Consulta Amigable)
        total_ejec = 0
        archivo = "%s-Seguimiento-PI.zip" % anio
        temporal = None
        try:
            with _abrir(URL_ARCHIVOS + archivo) as resp:
                conteo = _Conteo(resp, progreso("Ejecucion presupuestal (Consulta Amigable)"))
                temporal = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
                bloque = bytearray(1 << 20)
                while True:
                    n = conteo.readinto(bloque)
                    if not n:
                        break
                    temporal.write(bloque[:n])
                temporal.close()
            with zipfile.ZipFile(temporal.name) as z:
                nombre = [n for n in z.namelist() if n.lower().endswith(".csv")][0]
                with z.open(nombre) as crudo:
                    texto = io.TextIOWrapper(crudo, encoding="utf-8-sig", newline="")

                    def filtro_ejec(pos):
                        _exigir(pos, ["ANO_EJE", "FUNCION", "PRODUCTO_PROYECTO", "MONTO_PIM"], archivo)
                        a, f = pos["ANO_EJE"], pos["FUNCION"]
                        return lambda fila: fila[a].strip() == anio and fila[f].strip() == FUNCION_CODIGO

                    total_ejec, _ = _cargar(con, "ejecucion", csv.reader(texto), filtro_ejec,
                                            ["producto_proyecto"])
        finally:
            if temporal is not None:
                try:
                    os.unlink(temporal.name)
                except OSError:
                    pass

        ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        for k, v in (("actualizado", ahora), ("fecha_fuente", fecha_detalle), ("anio", anio),
                     ("funcion", FUNCION), ("inversiones", str(total_inv)),
                     ("f12b", str(total_f12b)), ("ejecucion", str(total_ejec))):
            con.execute("INSERT OR REPLACE INTO meta (k, v) VALUES (?, ?)", (k, v))
        avisar("Guardando", 1)
        con.commit()
        return {"inversiones": total_inv, "f12b": total_f12b, "ejecucion": total_ejec}
    except urllib.error.URLError as e:
        con.rollback()
        raise RuntimeError("No se pudo descargar del MEF. Revise su conexion a internet. (%s)" % e.reason)
    except BaseException:
        con.rollback()
        raise
    finally:
        con.close()


if __name__ == "__main__":
    def _mostrar(paso, avance):
        print("\r%-45s %3d%%" % (paso, avance * 100), end="", flush=True)
    print(actualizar(_mostrar))

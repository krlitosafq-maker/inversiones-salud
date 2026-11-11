"""Genera los archivos de datos que usa la app del celular (carpeta movil/datos)
a partir de la base de datos local. Se ejecuta despues de datos.py."""
import json
import os
import sys

import datos

DESTINO = os.path.join(datos.BASE, "movil", "datos")

INVERSIONES = [
    "clave", "codigo_unico", "codigo_snip", "nombre_inversion", "nivel", "sector", "entidad",
    "departamento", "provincia", "distrito", "tipo_inversion", "situacion", "marco",
    "registrado_pmi", "pmi_anio_1", "pmi_anio_2", "pmi_anio_3", "pmi_anio_4", "monto_viable",
    "costo_actualizado", "deven_acumul_anio_ant", "dev_anio_actual", "pia_anio_actual",
    "pim_anio_actual", "certif_anio_actual", "comprom_anual_anio_actual", "saldo_ejecutar",
    "avance_fisico", "avance_ejecucion", "primer_devengado", "ultimo_devengado", "nombre_opmi",
    "nombre_uf", "nombre_uei", "nombre_uep", "des_modalidad", "expediente_tecnico", "etapa_f8",
    "etapa_f9", "tiene_f12b", "fec_ini_ejecucion", "fec_fin_ejecucion", "fec_ini_ejec_fisica",
    "fec_fin_ejec_fisica", "fecha_viabilidad", "des_tipologia", "num_habitantes_benef",
]
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "set", "oct", "nov", "dic"]
F12B = ["dev_%s_ano_vig" % m for m in MESES] + [
    "fecha_ult_act_f12b", "ult_periodo_reg_f12b", "ult_estado_situacional", "ult_problema", "acc_problema"]
EJECUCION = ["ejecutora_nombre", "pliego_nombre", "fuente_financiamiento_nombre", "rubro_nombre"]
IDEAS = ["cod_idea", "nombre", "monto_inversion", "costo_estudio", "tipo", "nivel", "entidad",
         "fecha_registro", "estado", "anio", "departamento", "cod_unico"]


def _valor(v):
    """Acorta los valores para que el archivo pese menos."""
    if isinstance(v, float):
        return int(v) if v == int(v) or abs(v) >= 1000 else round(v, 2)
    if isinstance(v, str) and v.endswith(" 00:00:00"):
        return v[:10]
    return v


def _columnas(con, tabla, pedidas):
    hay = {f[1] for f in con.execute('PRAGMA table_info("%s")' % tabla)}
    return [c for c in pedidas if c in hay]


def _guardar(nombre, contenido):
    ruta = os.path.join(DESTINO, nombre)
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(contenido, f, ensure_ascii=False, separators=(",", ":"))
    return os.path.getsize(ruta)


def _tabla(con, tabla, pedidas, orden=""):
    cols = _columnas(con, tabla, pedidas)
    cur = con.execute('SELECT %s FROM "%s" %s' % (", ".join(cols), tabla, orden))
    return {"cols": cols, "filas": [[_valor(v) for v in f] for f in cur]}


def exportar():
    os.makedirs(DESTINO, exist_ok=True)
    con = datos.conectar()
    try:
        if not datos.hay_datos(con):
            raise RuntimeError("No hay datos. Ejecute primero datos.py")
        pesos = {}
        pesos["inversiones.json"] = _guardar("inversiones.json", _tabla(
            con, "inversiones", INVERSIONES, "ORDER BY pim_anio_actual DESC"))

        cols = _columnas(con, "f12b", F12B)
        f12b = {f[0]: [_valor(v) for v in f[1:]] for f in con.execute(
            "SELECT codigo_unico, %s FROM f12b" % ", ".join(cols))}
        pesos["f12b.json"] = _guardar("f12b.json", {"cols": cols, "filas": f12b})

        ejec = {}
        for f in con.execute(
                "SELECT producto_proyecto, %s, SUM(monto_pia), SUM(monto_pim), SUM(monto_devengado_ano_eje) "
                "FROM ejecucion GROUP BY 1, 2, 3, 4, 5 ORDER BY 7 DESC" % ", ".join(EJECUCION)):
            ejec.setdefault(f[0], []).append([_valor(v) for v in f[1:]])
        pesos["ejecucion.json"] = _guardar("ejecucion.json", {
            "cols": EJECUCION + ["pia", "pim", "devengado"], "filas": ejec})

        hay_ideas = con.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='ideas'").fetchone()
        meta = {f["k"]: f["v"] for f in con.execute("SELECT k, v FROM meta")}
        ruta_ideas = os.path.join(DESTINO, "ideas.json")
        if hay_ideas:
            ideas = _tabla(con, "ideas", IDEAS, "ORDER BY monto_inversion DESC")
            ideas["actualizado"] = meta.get("ideas_actualizado", "")
            pesos["ideas.json"] = _guardar("ideas.json", ideas)
        elif os.path.exists(ruta_ideas):
            # La consulta de ideas del MEF no responde desde fuera del Peru (caso de GitHub):
            # se conserva la ultima copia de ideas subida desde la computadora.
            pesos["ideas.json (copia anterior)"] = os.path.getsize(ruta_ideas)
        else:
            pesos["ideas.json"] = _guardar("ideas.json", {"cols": IDEAS, "filas": [], "actualizado": ""})

        _guardar("meta.json", meta)
        return pesos
    finally:
        con.close()


if __name__ == "__main__":
    for nombre, peso in exportar().items():
        print("%-18s %6.1f MB" % (nombre, peso / 1e6))
    sys.exit(0)

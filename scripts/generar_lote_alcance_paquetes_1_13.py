"""
Genera las 23 Actas de Alcance correspondientes a los Paquetes 1 al 13
(los que sí aparecen sincronizados en el OneDrive/SharePoint), a partir de
las actas iniciales ya localizadas.

REGLA DE NEGOCIO (instrucción de la Gerencia, relayada por el usuario):
va UNA Acta de Alcance por cada OFERENTE/inmueble, no una por paquete. En
Paquetes 1, 3, 4, 6 y 13 esto ya salía así de forma natural (cada subasta
ahí tiene un solo oferente). En Paquetes 9, 10, 11 y 12 el acta de Comité
cubre varios inmuebles/oferentes en una sola tabla, así que aquí se separa
esa tabla y se genera 1 Acta de Alcance por cada fila (ver
`_acta_por_inmueble`): 4 + 3 + 4 + 4 = 15 actas, más las 8 ya conformes
= 23 en total.

NOTA IMPORTANTE: el Código ActiBid, el oferente ganador y el monto
adjudicado (base de datos), y el valor catastral/PMV (Excel del paquete),
solo se conocen hoy para 1 de los oferentes de cada uno de los Paquetes 9,
10, 11 y 12 (el que ya se había consultado antes). Para los demás
oferentes de esos 4 paquetes esos campos quedan como "[POR COMPLETAR]"
hasta que se consiga esa información (base de datos o Excel); todo lo
demás (número/fecha de acta vieja, paquete, ID/FMI/Tipología/Dirección/
Municipio/Estado/Fecha subasta de CADA inmueble) sí sale completo de los
PDFs reales porque cada acta ya trae la tabla completa de todos ellos.

Casos especiales (Paquete 3 y Paquete 6 - Subasta 1.1): la versión
"vigente" que el usuario pidió usar (N17 / Ajustada) es un acta de SOLO
cambio de método de pago y no trae la tabla del inmueble, así que se
combina: la tabla de inmuebles sale del acta ORIGINAL (N16 / 019) y el
número/fecha de acta que se referencia en el Acta de Alcance es el de la
versión actualizada (No.17 / No.026), tal como pidió el usuario.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

PROYECTO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROYECTO / "scripts"))

from generar_acta_alcance import (
    leer_acta_inicial_archivo,
    generar_docx,
    _leer_consecutivo,
    _guardar_consecutivo,
    SALIDA,
)

BASE = Path(
    "/mnt/user-data/uploads/OneDrive - Activos por Colombia/Documentos/"
    "ActiBOX - Repositorio Documental - 2026_SOPORTES_SUBASTA"
)

BD_PENDIENTE = {
    "codigo_subasta": "—",
    "codigo_inmueble": "—",
    "nombre_ganador": "—",
    "id_ganador": "—",
    "monto_ganador": "0",
}

RUTA_DATOS_BD = Path(
    "/mnt/user-data/uploads/AUTOMATIZACION_WORDS/datos_alcance_paquetes_1_13.json"
)
DATOS_BD: dict = {}
if RUTA_DATOS_BD.is_file():
    DATOS_BD = json.loads(RUTA_DATOS_BD.read_text(encoding="utf-8"))

RUTA_DATOS_EXCEL = PROYECTO / "datos_excel_paquetes_1_13.json"
DATOS_EXCEL: dict = {}
if RUTA_DATOS_EXCEL.is_file():
    DATOS_EXCEL = json.loads(RUTA_DATOS_EXCEL.read_text(encoding="utf-8"))


def _excel_para(identificador: str) -> dict:
    entrada = DATOS_EXCEL.get(identificador)
    if entrada and entrada.get("ok"):
        return entrada
    return {}


def _bd_para(identificador: str) -> dict:
    entrada = DATOS_BD.get(identificador)
    if entrada and entrada.get("ok"):
        return {
            "codigo_subasta": entrada.get("codigo_subasta", "—"),
            "codigo_inmueble": entrada.get("codigo_inmueble", "—"),
            "nombre_ganador": entrada.get("nombre_ganador", "—"),
            "id_ganador": entrada.get("id_ganador", "—"),
            "monto_ganador": entrada.get("monto_ganador", "0"),
        }
    return BD_PENDIENTE


def _acta_simple(ruta_rel: str, identificador_archivo: str, numero_paquete_forzado: str | None = None) -> tuple[dict, str]:
    ruta = BASE / ruta_rel
    acta = leer_acta_inicial_archivo(ruta)
    if numero_paquete_forzado and acta.get("numero_paquete") == "—":
        acta["numero_paquete"] = numero_paquete_forzado
    return acta, identificador_archivo


def _acta_combinada(ruta_original_rel: str, ruta_actualizada_rel: str, identificador_archivo: str,
                     numero_paquete_forzado: str | None = None) -> tuple[dict, str]:
    """Tabla de inmuebles del acta ORIGINAL + número/fecha del acta ACTUALIZADA."""
    original = leer_acta_inicial_archivo(BASE / ruta_original_rel)
    actualizada = leer_acta_inicial_archivo(BASE / ruta_actualizada_rel)
    combinada = dict(original)
    combinada["numero_acta_vieja"] = actualizada["numero_acta_vieja"]
    combinada["fecha_acta_vieja"] = actualizada["fecha_acta_vieja"]
    if numero_paquete_forzado and combinada.get("numero_paquete") == "—":
        combinada["numero_paquete"] = numero_paquete_forzado
    return combinada, identificador_archivo


def _acta_por_inmueble(ruta_rel: str, fmi_filtro: str, identificador_archivo: str,
                        numero_paquete_forzado: str | None = None) -> tuple[dict, str]:
    """Igual que _acta_simple, pero cuando el acta vieja trae VARIOS inmuebles
    en una sola tabla (Paquetes 9, 10, 11 y 12: un acta de comité por todo el
    paquete, con un oferente distinto por cada inmueble), se queda SOLO con
    la fila del inmueble indicado (fmi_filtro) -- así cada oferente recibe su
    propia Acta de Alcance individual, en vez de una sola con todos juntos."""
    ruta = BASE / ruta_rel
    acta = leer_acta_inicial_archivo(ruta)
    inmueble = next((i for i in acta["inmuebles"] if i["fmi"] == fmi_filtro), None)
    if inmueble is None:
        disponibles = [i["fmi"] for i in acta["inmuebles"]]
        raise ValueError(
            f"No se encontró el FMI '{fmi_filtro}' en {ruta_rel}. Disponibles: {disponibles}"
        )
    acta = dict(acta)
    acta["inmuebles"] = [inmueble]
    if numero_paquete_forzado and acta.get("numero_paquete") == "—":
        acta["numero_paquete"] = numero_paquete_forzado
    return acta, identificador_archivo


CASOS = [
    ("Paquete 1 - Subasta 1.1 (Acta 015)", *_acta_simple(
        "PAQUETE_1/1.1_SUBASTA_1771026814299217686_(FMI 50C-814138)/1.1.1_IMPORTADORA_DE_INSERTOS_SAS/"
        "1.1.1.4_COMERCIALIZACION/1.1.1.4.2 ACTAS/005_FORMATO_ACTA_DE_COMITE_SUBASTAS_015_13032026 (1) (1).pdf",
        "50C-814138")),
    ("Paquete 1 - Subasta 1.2 (Acta 014)", *_acta_simple(
        "PAQUETE_1/1.2_SUBASTA_1771026813588996293_(FMI 200-179598)/1.2.1_MARIO_GARZON/"
        "1.2.1.4 COMERCIALIZACION/1.2.1.4.2_ACTAS/005_FORMATO_ACTA_DE_COMITE_SUBASTAS_014_13032026.pdf",
        "200-179598")),
    ("Paquete 1 - Subasta 1.3 (Acta 013, referenciada como No.071)", *_acta_combinada(
        "PAQUETE_1/1.3_SUBASTA_1771026813306252156_(FMI 50C-1257163)/1.3.1_DANIEL_CASTRO_CORREA_GANADOR/"
        "1.3.1.4_COMERCIALIZACION/1.3.1.4.2_ACTAS/005_FORMATO_ACTA_DE_COMITE_SUBASTAS_013_13032026 (1).pdf",
        "PAQUETE_1/1.3_SUBASTA_1771026813306252156_(FMI 50C-1257163)/1.3.1_DANIEL_CASTRO_CORREA_GANADOR/"
        "1.3.1.4_COMERCIALIZACION/1.3.1.4.2_ACTAS/FORMATO_ACTA_DE_COMITE_071_cambio_metodo_de_pago.pdf",
        "50C-1257163")),
    ("Paquete 3 (Acta 16, referenciada como No.17)", *_acta_combinada(
        "PAQUETE_3/1.1_SUBASTA_1771943132210948434_(FMI 060-279174)/1.1.1_ GLADYS_EDITH_MACHADO_ACUÑA/"
        "1.1.1.4_COMERCIALIZACION/1.1.1.4.2_ACTAS/005_ACTA_DE_COMITE_DE_ADJUDICACION_N16.pdf",
        "PAQUETE_3/1.1_SUBASTA_1771943132210948434_(FMI 060-279174)/1.1.1_ GLADYS_EDITH_MACHADO_ACUÑA/"
        "1.1.1.4_COMERCIALIZACION/1.1.1.4.2_ACTAS/005_ACTA_DE_COMITE_DE_ADJUDICACION_N17_CAMBIO_METODO_DE_PAGO.pdf",
        "060-279174")),
    ("Paquete 4 (Acta 018)", *_acta_simple(
        "PAQUETE_4/1.1_SUBASTA_20260407150149800_(FMI 001-313399)/1.1.1_ SEBASTIAN_MESA_PEREZ/"
        "1.1.1.4_COMERCIALIZACION/1.1.1.4.2_ACTAS/005_FORMATO_ACTA_DE_COMITE_SUBASTAS_018_07042026_sebastian mesa.pdf",
        "001-313399")),
    ("Paquete 6 - Subasta 1.1 (Acta 19, referenciada como No.026)", *_acta_combinada(
        "PAQUETE_6/1.1_SUBASTA_20260408150353250_(FMI 001-790674,001-790636,001-790637,001-790580,001-790582)/"
        "1.1.1_SEBASTIAN_BUILES_VARGAS/1.1.1.4_COMERCIALIZACION/1.1.1.4.2_ACTAS/"
        "005_FORMATO_ACTA_DE_COMITE_SUBASTAS_019_08042026.pdf",
        "PAQUETE_6/1.1_SUBASTA_20260408150353250_(FMI 001-790674,001-790636,001-790637,001-790580,001-790582)/"
        "1.1.1_SEBASTIAN_BUILES_VARGAS/1.1.1.4_COMERCIALIZACION/1.1.1.4.2_ACTAS/"
        "005_FORMATO_ACTA_DE_COMITE_SUBASTAS_SEBASTIAN_BUILES_AJUSTADA.pdf",
        "001-790674")),
    ("Paquete 6 - Subasta 1.2 (Acta 020)", *_acta_simple(
        "PAQUETE_6/1.2_SUBASTA_20260408150058671_(FMI 50C-332612)/1.2.1_ALBERTO_ARBIL_Y_WILSON_GAITAN/"
        "1.2.1.4_COMERCIALIZACION/1.2.1.4.2_ACTAS/005_FORMATO_ACTA_DE_COMITE_SUBASTAS_020_09042026_Albertoabril.pdf",
        "50C-332612")),
    # Paquete 9 (Acta 023): 4 inmuebles / 4 oferentes -> 1 Acta de Alcance por cada uno.
    ("Paquete 9 - Oferente 1 de 4 (Rocío Jaramillo, FMI 001-42028)", *_acta_por_inmueble(
        "PAQUETE_9/FORMATO_ACTA_DE_COMITE_SUBASTAS_023.pdf", "001-42028", "001-42028")),
    ("Paquete 9 - Oferente 2 de 4 (FMI 060-91447)", *_acta_por_inmueble(
        "PAQUETE_9/FORMATO_ACTA_DE_COMITE_SUBASTAS_023.pdf", "060-91447", "060-91447")),
    ("Paquete 9 - Oferente 3 de 4 (FMI 001-868645)", *_acta_por_inmueble(
        "PAQUETE_9/FORMATO_ACTA_DE_COMITE_SUBASTAS_023.pdf", "001-868645", "001-868645")),
    ("Paquete 9 - Oferente 4 de 4 (FMI 280-39702/280-39703)", *_acta_por_inmueble(
        "PAQUETE_9/FORMATO_ACTA_DE_COMITE_SUBASTAS_023.pdf", "280-39702/280-39703", "280-39702_280-39703")),

    # Paquete 10 (Acta 22): 3 inmuebles / 3 oferentes -> 1 Acta de Alcance por cada uno.
    ("Paquete 10 - Oferente 1 de 3 (FMI 200-88115)", *_acta_por_inmueble(
        "PAQUETE_10/FORMATO_ACTA_DE_COMITE_22.pdf", "200-88115", "200-88115", numero_paquete_forzado="10")),
    ("Paquete 10 - Oferente 2 de 3 (FMI 200-213538)", *_acta_por_inmueble(
        "PAQUETE_10/FORMATO_ACTA_DE_COMITE_22.pdf", "200-213538", "200-213538", numero_paquete_forzado="10")),
    ("Paquete 10 - Oferente 3 de 3 (FMI 378-141134/378-140743/378-140746/378-140747/378-140748)", *_acta_por_inmueble(
        "PAQUETE_10/FORMATO_ACTA_DE_COMITE_22.pdf", "378-141134378-140743378-140746378-140747378-140748",
        "378-141134_378-140743_378-140746_378-140747_378-140748", numero_paquete_forzado="10")),

    # Paquete 11 (Acta 024): 4 inmuebles / 4 oferentes -> 1 Acta de Alcance por cada uno.
    ("Paquete 11 - Oferente 1 de 4 (FMI 001-977915/001-977935/001-977931)", *_acta_por_inmueble(
        "PAQUETE_11/FORMATO_ACTA_DE_COMITE_SUBASTAS_024.pdf", "001-977915001-977935001-977931",
        "001-977915_001-977935_001-977931")),
    ("Paquete 11 - Oferente 2 de 4 (Gisella Bernal, FMI 50N-1112089, terminación anticipada)", *_acta_por_inmueble(
        "PAQUETE_11/FORMATO_ACTA_DE_COMITE_SUBASTAS_024.pdf", "50N-1112089", "50N-1112089")),
    ("Paquete 11 - Oferente 3 de 4 (FMI 280-150766)", *_acta_por_inmueble(
        "PAQUETE_11/FORMATO_ACTA_DE_COMITE_SUBASTAS_024.pdf", "280-150766", "280-150766")),
    ("Paquete 11 - Oferente 4 de 4 (FMI 350-123639)", *_acta_por_inmueble(
        "PAQUETE_11/FORMATO_ACTA_DE_COMITE_SUBASTAS_024.pdf", "350-123639", "350-123639")),

    # Paquete 12 (Acta Comité Técnico Sociedades): 4 inmuebles / 4 oferentes -> 1 Acta de Alcance por cada uno.
    ("Paquete 12 - Oferente 1 de 4 (Cindy Campo, FMI 01N-5168178)", *_acta_por_inmueble(
        "PAQUETE_12/1.1_SUBASTA_(FMI_01N-5168178)/1.1.1_CINDY_KATHERINE_CAMPO_USUGA/"
        "1.1.1.4_COMERCIALIZACION/1.1.1.4.2_ACTAS/005_ACTA_DEL_COMITE_TECNICO_DE_SOCIEDADES.pdf",
        "01N-5168178", "01N-5168178")),
    ("Paquete 12 - Oferente 2 de 4 (FMI 230-127895)", *_acta_por_inmueble(
        "PAQUETE_12/1.1_SUBASTA_(FMI_01N-5168178)/1.1.1_CINDY_KATHERINE_CAMPO_USUGA/"
        "1.1.1.4_COMERCIALIZACION/1.1.1.4.2_ACTAS/005_ACTA_DEL_COMITE_TECNICO_DE_SOCIEDADES.pdf",
        "230-127895", "230-127895")),
    ("Paquete 12 - Oferente 3 de 4 (FMI 50C-159617)", *_acta_por_inmueble(
        "PAQUETE_12/1.1_SUBASTA_(FMI_01N-5168178)/1.1.1_CINDY_KATHERINE_CAMPO_USUGA/"
        "1.1.1.4_COMERCIALIZACION/1.1.1.4.2_ACTAS/005_ACTA_DEL_COMITE_TECNICO_DE_SOCIEDADES.pdf",
        "50C-159617", "50C-159617")),
    ("Paquete 12 - Oferente 4 de 4 (FMI 50N-20394782/.../50N-20394835)", *_acta_por_inmueble(
        "PAQUETE_12/1.1_SUBASTA_(FMI_01N-5168178)/1.1.1_CINDY_KATHERINE_CAMPO_USUGA/"
        "1.1.1.4_COMERCIALIZACION/1.1.1.4.2_ACTAS/005_ACTA_DEL_COMITE_TECNICO_DE_SOCIEDADES.pdf",
        "50N-20394782/50N-20394813/50N-20394814/50N-20394835",
        "50N-20394782_50N-20394813_50N-20394814_50N-20394835")),
    ("Paquete 13 (Acta 025)", *_acta_simple(
        "PAQUETE_13/FORMATO_ACTA_DE_COMITE_SUBASTAS_024 PAQUETE 13.pdf", "001-911997")),
]

# Casos donde el Comité de Adjudicación declaró la TERMINACIÓN ANTICIPADA del
# negocio (adjudicación dejada sin efecto) mediante un acta posterior. Para
# estos, el "oferente ganador" que devolvería la base de datos ya no aplica:
# se reemplaza por la nota de terminación, dejando el monto originalmente
# ofertado como dato histórico (así el cálculo de "¿superó el precio base?"
# sigue reflejando la realidad: no lo superó, por eso se terminó el proceso).
TERMINACIONES_ANTICIPADAS = {
    "50N-1112089": {
        "nota": "SIN EFECTO (Terminación anticipada, Acta de Comité No. 50 del "
                "22/07/2026): valor adjudicado a Gisella P. Bernal G. inferior al "
                "avalúo catastral 2026; incumple la Política de PMV de la SAE.",
        "id": "N/A",
    },
}

# Fuerza el consecutivo a partir de 81 para esta corrida (independiente de
# lo que haya quedado guardado de pruebas anteriores).
consecutivo = 81
resumen = []

for nombre, acta_vieja, identificador in CASOS:
    codigo_archivo = identificador.replace("/", "_").replace(" ", "_")
    ruta_salida = SALIDA / f"ACTA_DE_ALCANCE_{consecutivo}_{codigo_archivo}.docx"
    bd = _bd_para(identificador)
    if identificador in TERMINACIONES_ANTICIPADAS:
        terminacion = TERMINACIONES_ANTICIPADAS[identificador]
        bd = dict(bd)
        bd["nombre_ganador"] = terminacion["nota"]
        bd["id_ganador"] = terminacion["id"]
    excel = _excel_para(identificador)
    pendientes = generar_docx(acta_vieja, bd, consecutivo, ruta_salida, excel=excel)
    resumen.append((consecutivo, nombre, acta_vieja["numero_acta_vieja"], len(acta_vieja["inmuebles"]), ruta_salida))
    print(f"[{consecutivo}] {nombre} -> {ruta_salida.name}  (acta vieja No.{acta_vieja['numero_acta_vieja']}, "
          f"{len(acta_vieja['inmuebles'])} inmueble(s))")
    consecutivo += 1

_guardar_consecutivo(consecutivo)
print(f"\nListo. Próximo consecutivo disponible: {consecutivo}")

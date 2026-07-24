"""Lógica compartida: conexión (local.conf / .env) y consulta informe Polybid."""

from __future__ import annotations

import base64
import html as html_module
import json
import os
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus
from uuid import UUID


def get_repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


def load_dotenv_files() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    root = get_repo_root()
    load_dotenv(root / ".env")
    load_dotenv(Path(__file__).resolve().parent / ".env")
    load_dotenv()


def json_default(obj: Any) -> Any:
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    if isinstance(obj, Decimal):
        return str(obj)
    if isinstance(obj, UUID):
        return str(obj)
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")


def dsn_from_local_conf(conf_path: Path) -> str | None:
    if not conf_path.is_file():
        return None
    try:
        with open(conf_path, encoding="utf-8") as f:
            cfg = json.load(f)
    except (OSError, json.JSONDecodeError):
        return None
    db = cfg.get("database")
    if not isinstance(db, dict):
        return None
    user = (db.get("user") or "").strip()
    password = "" if db.get("password") is None else str(db.get("password"))
    database = (db.get("database") or "").strip()
    if not user or not database:
        return None
    address = (db.get("address") or "127.0.0.1:5432").strip()
    if ":" in address:
        host, port = address.rsplit(":", 1)
        host, port = host.strip(), port.strip()
    else:
        host, port = address, "5432"
    params = (db.get("parameters") or "").strip().lstrip("?&")
    userinfo = f"{quote_plus(user)}:{quote_plus(password)}"
    url = f"postgresql://{userinfo}@{host}:{port}/{quote_plus(database)}"
    if params:
        url = f"{url}?{params}"
    return url


def resolve_local_conf_path(explicit: str | None) -> Path:
    if explicit:
        p = Path(explicit).expanduser()
        return p if p.is_absolute() else (get_repo_root() / p).resolve()
    file_conf = os.environ.get("FILE_CONF", "local.conf").strip() or "local.conf"
    path_conf = os.environ.get("PATH_CONF", "").strip()
    root = get_repo_root()
    if path_conf:
        base = Path(path_conf)
        if not base.is_absolute():
            for candidate in (
                (Path.cwd() / base).resolve(),
                (root / base).resolve(),
            ):
                if candidate.is_dir() and (candidate / file_conf).is_file():
                    return candidate / file_conf
                if candidate.is_file():
                    return candidate
        elif base.is_dir() and (base / file_conf).is_file():
            return base / file_conf
        elif base.is_file():
            return base
    return root / "config" / file_conf


def _dsn_from_env_parts() -> str | None:
    user = (os.environ.get("PGUSER") or os.environ.get("POSTGRES_USER") or "").strip()
    if os.environ.get("PGPASSWORD") is not None:
        password = os.environ.get("PGPASSWORD", "")
    else:
        password = os.environ.get("POSTGRES_PASSWORD", "")
    dbname = (os.environ.get("PGDATABASE") or os.environ.get("POSTGRES_DB") or "").strip()
    if not user or not dbname:
        return None
    host = (os.environ.get("PGHOST") or os.environ.get("POSTGRES_HOST") or "localhost").strip()
    port = (os.environ.get("PGPORT") or os.environ.get("POSTGRES_PORT") or "5432").strip()
    return (
        "postgresql://"
        f"{quote_plus(user)}:{quote_plus(str(password))}@{quote_plus(host)}:{port}/{quote_plus(dbname)}"
    )


def get_connection(dsn_override: str | None, local_conf: Path) -> Any:
    try:
        import psycopg
    except ImportError as e:
        raise ImportError(
            "Falta psycopg. Instala: pip install -r informe_subasta_polybid/requirements.txt"
        ) from e

    if dsn_override and dsn_override.strip():
        return psycopg.connect(dsn_override.strip())

    dsn = os.environ.get("DATABASE_URL", "").strip()
    if not dsn:
        dsn = dsn_from_local_conf(local_conf) or ""
    if not dsn:
        dsn = _dsn_from_env_parts() or ""
    if dsn:
        return psycopg.connect(dsn)

    kwargs = {
        "host": os.environ.get("PGHOST", os.environ.get("POSTGRES_HOST", "localhost")),
        "port": os.environ.get("PGPORT", os.environ.get("POSTGRES_PORT", "5432")),
        "dbname": os.environ.get("PGDATABASE", os.environ.get("POSTGRES_DB", "postgres")),
        "user": os.environ.get(
            "PGUSER",
            os.environ.get("POSTGRES_USER", os.environ.get("USER", "postgres")),
        ),
    }
    pwd = os.environ.get("PGPASSWORD")
    if pwd is None:
        pwd = os.environ.get("POSTGRES_PASSWORD")
    if pwd is not None:
        kwargs["password"] = pwd

    if local_conf.is_file():
        try:
            with open(local_conf, encoding="utf-8") as f:
                cfg = json.load(f)
            db = cfg.get("database") or {}
            addr = (db.get("address") or "").strip()
            if addr and ":" in addr:
                h, p = addr.rsplit(":", 1)
                kwargs["host"] = h.strip()
                kwargs["port"] = p.strip()
            elif addr:
                kwargs["host"] = addr
            if db.get("user"):
                kwargs["user"] = str(db["user"])
            if db.get("database"):
                kwargs["dbname"] = str(db["database"])
            if db.get("password") is not None:
                kwargs["password"] = str(db["password"])
        except (OSError, json.JSONDecodeError, KeyError):
            pass

    return psycopg.connect(**kwargs)


def fetch_informe(conn: Any, auction_id: str) -> dict[str, Any]:
    UUID(auction_id)

    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT id, code, title, status, initial_value, start_date, end_date
            FROM polybid.auctions
            WHERE id = %s::uuid
            """,
            (auction_id,),
        )
        row = cur.fetchone()
        if not row:
            raise ValueError(f"No existe subasta con id={auction_id}")

        cols = [d.name for d in cur.description]
        subasta = dict(zip(cols, row))

        cur.execute(
            """
            SELECT COUNT(*)::int
            FROM polybid.auction_bids
            WHERE auction_id = %s::uuid
            """,
            (auction_id,),
        )
        (total_pujas,) = cur.fetchone()

        cur.execute(
            """
            SELECT
              p.id,
              p.client_id,
              p.display_name,
              p.status,
              p.created_at,
              p.approved_at,
              ct.id AS contact_tercero_id,
              ct.nombre_principal,
              ct.email
            FROM polybid.auction_participants p
            LEFT JOIN polibid_credentials pc ON pc.client_id = p.client_id
            LEFT JOIN contact_terceros ct ON ct.id = pc.contact_tercero_id
            WHERE p.auction_id = %s::uuid
            ORDER BY p.created_at ASC
            """,
            (auction_id,),
        )
        pcols = [d.name for d in cur.description]
        participantes = [dict(zip(pcols, r)) for r in cur.fetchall()]

        sql_ganador = """
            SELECT
              b.id AS bid_id,
              b.amount,
              b.status,
              b.created_at,
              p.id AS participant_id,
              p.client_id,
              p.display_name,
              ct.nombre_principal,
              ct.email
            FROM polybid.auction_bids b
            INNER JOIN polybid.auction_participants p
              ON p.auction_id = b.auction_id AND p.client_id = b.client_id
            LEFT JOIN polibid_credentials pc ON pc.client_id = b.client_id
            LEFT JOIN contact_terceros ct ON ct.id = pc.contact_tercero_id
            WHERE b.auction_id = %s::uuid AND b.status = 'WINNING'
            LIMIT 1
            """
        cur.execute(sql_ganador, (auction_id,))
        wcols = [d.name for d in cur.description]
        wrow = cur.fetchone()
        informe_mejor_oferta = False
        if not wrow:
            # Subastas ACTIVE: a veces aún no existe fila WINNING; se toma la mejor oferta vigente.
            cur.execute(
                """
                SELECT
                  b.id AS bid_id,
                  b.amount,
                  b.status,
                  b.created_at,
                  p.id AS participant_id,
                  p.client_id,
                  p.display_name,
                  ct.nombre_principal,
                  ct.email
                FROM polybid.auction_bids b
                INNER JOIN polybid.auction_participants p
                  ON p.auction_id = b.auction_id AND p.client_id = b.client_id
                LEFT JOIN polibid_credentials pc ON pc.client_id = b.client_id
                LEFT JOIN contact_terceros ct ON ct.id = pc.contact_tercero_id
                WHERE b.auction_id = %s::uuid
                ORDER BY b.amount DESC, b.created_at DESC
                LIMIT 1
                """,
                (auction_id,),
            )
            wrow = cur.fetchone()
            informe_mejor_oferta = bool(wrow)

        ganador = dict(zip(wcols, wrow)) if wrow else None
        if ganador is not None:
            ganador["informe_mejor_oferta"] = informe_mejor_oferta

        cur.execute(
            """
            SELECT
              b.id AS bid_id,
              b.amount,
              b.status,
              b.created_at,
              b.ip_address::text AS ip_address,
              b.user_agent,
              ct.nombre_principal,
              ct.email
            FROM polybid.auction_bids b
            LEFT JOIN polybid.auction_participants p
              ON p.auction_id = b.auction_id AND p.client_id = b.client_id
            LEFT JOIN polibid_credentials pc ON pc.client_id = b.client_id
            LEFT JOIN contact_terceros ct ON ct.id = pc.contact_tercero_id
            WHERE b.auction_id = %s::uuid
            ORDER BY b.created_at ASC
            """,
            (auction_id,),
        )
        hcols = [d.name for d in cur.description]
        historial_pujas = [
            _normalize_historial_puja_row(dict(zip(hcols, row))) for row in cur.fetchall()
        ]

        # ActiBID: contacto / inmueble / manifestaciones (solo para vista previa en app, no va a Teams)
        vista_previa_extra: dict[str, Any] = {
            "contacto": None,
            "inmueble": None,
            "manifestacion_interes": [],
            "manifestacion_workflow": [],
        }
        cur.execute(
            """
            SELECT contact_tercero_id, inmueble_id
            FROM polibid_subastas_v2
            WHERE auction_id = %s::uuid
            ORDER BY id DESC
            LIMIT 1
            """,
            (auction_id,),
        )
        link_row = cur.fetchone()
        if link_row:
            ct_id, inm_id = link_row[0], link_row[1]
            if ct_id:
                cur.execute(
                    """
                    SELECT id, tipo_tercero, tipo_solicitud, identificacion_tipo, identificacion_numero,
                           nombre_principal, direccion_principal, ciudad, departamento, pais,
                           email, celular, telefono_fijo, fecha_diligenciamiento
                    FROM contact_terceros
                    WHERE id = %s
                    """,
                    (ct_id,),
                )
                r_ct = cur.fetchone()
                if r_ct:
                    ccols = [d.name for d in cur.description]
                    vista_previa_extra["contacto"] = dict(zip(ccols, r_ct))
            if inm_id:
                cur.execute(
                    """
                    SELECT id, codigo, referencia, numero_matricula, estrato,
                           area_construida, area_privada, area_lote, habitaciones, banos, parqueaderos,
                           precio_base_venta, valor_minimo_venta, valor_avaluo_comercial, valor_avaluo_catastral,
                           gestion, fecha_cambio_estado
                    FROM mst_inmuebles
                    WHERE id = %s
                    """,
                    (inm_id,),
                )
                r_in = cur.fetchone()
                if r_in:
                    icols = [d.name for d in cur.description]
                    vista_previa_extra["inmueble"] = dict(zip(icols, r_in))
                cur.execute(
                    """
                    SELECT id, estado, observaciones, situacion_laboral, created_at, updated_at
                    FROM manifestacion_interes
                    WHERE inmueble_id = %s
                    ORDER BY updated_at DESC NULLS LAST
                    LIMIT 20
                    """,
                    (inm_id,),
                )
                micols = [d.name for d in cur.description]
                vista_previa_extra["manifestacion_interes"] = [
                    dict(zip(micols, row)) for row in cur.fetchall()
                ]
                cur.execute(
                    """
                    SELECT id, client_id, status, current_step, notes, created_at, updated_at
                    FROM mst_interest_manifestations
                    WHERE property_id = %s
                    ORDER BY created_at DESC
                    LIMIT 20
                    """,
                    (inm_id,),
                )
                mwcols = [d.name for d in cur.description]
                vista_previa_extra["manifestacion_workflow"] = [
                    dict(zip(mwcols, row)) for row in cur.fetchall()
                ]

    inscritos = len(participantes)
    return {
        "auction_id": auction_id,
        "subasta": subasta,
        "resumen": {
            "participantes_inscritos": inscritos,
            "participantes_con_ofertas_numero_pujas": total_pujas,
            "total_ofertas_recibidas": total_pujas,
        },
        "participantes": participantes,
        "ganador": ganador,
        "historial_pujas": historial_pujas,
        "vista_previa_extra": vista_previa_extra,
    }


def _esc(s: Any) -> str:
    return html_module.escape(str(s if s is not None else ""))


def _html_kv_table(title: str, data: dict[str, Any] | None, labels: dict[str, str]) -> str:
    if not data:
        return ""
    rows = []
    for k, lab in labels.items():
        if k not in data:
            continue
        v = data.get(k)
        if v is None or v == "":
            continue
        rows.append(
            f"<tr><td style=\"border:1px solid #c8c6c4;padding:8px 12px;background:#f3f2f1;width:32%;"
            f"font-weight:600;\">{_esc(lab)}</td>"
            f"<td style=\"border:1px solid #c8c6c4;padding:8px 12px;\">{_esc(v)}</td></tr>"
        )
    if not rows:
        return ""
    return (
        f"<h4 style=\"margin:20px 0 10px 0;color:#201f1e;font-size:15px;\">{_esc(title)}</h4>"
        f"<table style=\"border-collapse:collapse;width:100%;max-width:920px;\">"
        f"{''.join(rows)}</table>"
    )


def vista_previa_extra_html(extra: dict[str, Any]) -> str:
    """Bloque HTML solo para la vista previa en Streamlit (contacto, inmueble, manifestaciones)."""
    contacto = extra.get("contacto")
    inmueble = extra.get("inmueble")
    m_interes = extra.get("manifestacion_interes") or []
    m_wf = extra.get("manifestacion_workflow") or []

    if not contacto and not inmueble and not m_interes and not m_wf:
        return (
            "<p style=\"color:#605e5c;font-size:13px;margin-top:12px;\">"
            "No hay registro en ActiBID vinculado a esta subasta en <code>polibid_subastas_v2</code> "
            "(contacto / inmueble), o aún no hay manifestación para ese inmueble."
            "</p>"
        )

    parts: list[str] = ['<div style="margin-top:20px;padding-top:16px;border-top:2px solid #edebe9;">']
    parts.append(
        "<h3 style=\"margin:0 0 12px 0;color:#323130;font-size:17px;\">Más información (solo vista previa)</h3>"
    )

    parts.append(
        _html_kv_table(
            "Contacto (tercero vinculado a la subasta)",
            contacto,
            {
                "id": "ID contacto",
                "nombre_principal": "Nombre",
                "identificacion_tipo": "Tipo identificación",
                "identificacion_numero": "Número identificación",
                "tipo_tercero": "Tipo tercero",
                "tipo_solicitud": "Tipo solicitud",
                "email": "Correo",
                "celular": "Celular",
                "telefono_fijo": "Teléfono fijo",
                "direccion_principal": "Dirección",
                "ciudad": "Ciudad",
                "departamento": "Departamento",
                "pais": "País",
                "fecha_diligenciamiento": "Fecha diligenciamiento",
            },
        )
    )

    parts.append(
        _html_kv_table(
            "Inmueble",
            inmueble,
            {
                "id": "ID inmueble",
                "codigo": "Código",
                "referencia": "Referencia",
                "numero_matricula": "Matrícula (FMI)",
                "estrato": "Estrato",
                "area_lote": "Área lote",
                "area_construida": "Área construida",
                "area_privada": "Área privada",
                "habitaciones": "Habitaciones",
                "banos": "Baños",
                "parqueaderos": "Parqueaderos",
                "precio_base_venta": "Precio base venta",
                "valor_minimo_venta": "Valor mínimo venta",
                "valor_avaluo_comercial": "Avalúo comercial",
                "valor_avaluo_catastral": "Avalúo catastral",
                "gestion": "Gestión",
                "fecha_cambio_estado": "Fecha cambio estado",
            },
        )
    )

    if m_interes:
        parts.append(
            "<h4 style=\"margin:20px 0 10px 0;color:#201f1e;font-size:15px;\">"
            "Manifestación de interés</h4>"
        )
        parts.append(
            "<table style=\"border-collapse:collapse;width:100%;max-width:920px;font-size:13px;\">"
            "<thead><tr>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">ID</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Estado</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Situación laboral</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Observaciones</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Alta</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Actualización</th>"
            "</tr></thead><tbody>"
        )
        for row in m_interes:
            parts.append(
                "<tr>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('id'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('estado'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('situacion_laboral'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('observaciones'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('created_at'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('updated_at'))}</td>"
                "</tr>"
            )
        parts.append("</tbody></table>")

    if m_wf:
        parts.append(
            "<h4 style=\"margin:20px 0 10px 0;color:#201f1e;font-size:15px;\">"
            "Manifestación (workflow / pasos)</h4>"
        )
        parts.append(
            "<table style=\"border-collapse:collapse;width:100%;max-width:920px;font-size:13px;\">"
            "<thead><tr>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">ID</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Cliente</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Estado</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Paso actual</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Notas</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Alta</th>"
            "<th style=\"border:1px solid #c8c6c4;padding:8px;background:#edebe9;\">Actualización</th>"
            "</tr></thead><tbody>"
        )
        for row in m_wf:
            parts.append(
                "<tr>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('id'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('client_id'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('status'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('current_step'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('notes'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('created_at'))}</td>"
                f"<td style=\"border:1px solid #c8c6c4;padding:8px;\">{_esc(row.get('updated_at'))}</td>"
                "</tr>"
            )
        parts.append("</tbody></table>")

    parts.append("</div>")
    return "".join(parts)


def _fmt_dt(val: Any) -> str:
    if val is None:
        return "—"
    s = str(val)
    return s.split(".")[0] if "." in s else s


def _format_cop_co(amount: Any) -> str:
    try:
        from decimal import Decimal

        d = Decimal(str(amount).replace(",", "").replace(" ", ""))
    except Exception:
        return str(amount)
    s = f"{d:.2f}"
    intpart, _, frac = s.partition(".")
    frac = (frac + "00")[:2]
    neg = intpart.startswith("-")
    if neg:
        intpart = intpart[1:]
    rev = intpart[::-1]
    chunks = [rev[i : i + 3] for i in range(0, len(rev), 3)]
    int_fmt = ".".join(c[::-1] for c in reversed(chunks))
    if neg:
        int_fmt = "-" + int_fmt
    return f"$ {int_fmt},{frac}"


def _truncate_text(s: Any, max_len: int = 56) -> str:
    t = (str(s) if s is not None else "").strip()
    if not t:
        return "—"
    if len(t) <= max_len:
        return t
    return t[: max_len - 1] + "…"


def _normalize_historial_puja_row(r: dict[str, Any]) -> dict[str, Any]:
    """Fila lista para informe: usuario, monto, tipo (recorte user_agent), ip, fecha, estado."""
    nombre = (r.get("nombre_principal") or "").strip()
    email = (r.get("email") or "").strip()
    if nombre or email:
        usuario = f"{nombre or '—'} · {email or '—'}"
    else:
        usuario = "—"
    ip = (r.get("ip_address") or "").strip() or "—"
    return {
        "bid_id": r.get("bid_id"),
        "usuario": usuario,
        "monto": r.get("amount"),
        "tipo": _truncate_text(r.get("user_agent"), 56),
        "ip": ip,
        "fecha": r.get("created_at"),
        "estado": (r.get("status") or "").strip() or "—",
    }


def informe_teams_plain(data: dict[str, Any]) -> str:
    """Texto plano con tabuladores (respaldo si el HTML no pega)."""
    lines: list[str] = []
    r, s = data["resumen"], data["subasta"]
    lines.append("INFORME DE SUBASTA")
    lines.append(f"Identificador:\t{data['auction_id']}")
    lines.append("")
    lines.append("Subasta")
    lines.append(f"Código:\t{s.get('code')}")
    lines.append(f"Estado:\t{s.get('status')}")
    lines.append(f"Título:\t{s.get('title')}")
    lines.append("")
    lines.append("Resumen")
    lines.append("Métrica\tValor")
    lines.append(f"Participantes inscritos\t{r['participantes_inscritos']}")
    lines.append(
        "Participantes con ofertas (número de pujas registradas)\t"
        f"{r['participantes_con_ofertas_numero_pujas']}"
    )
    lines.append(f"Total de ofertas recibidas\t{r['total_ofertas_recibidas']}")
    lines.append("")
    lines.append("Detalle por participante")
    lines.append(
        "ID participante\tUsuario (nombre · correo)\tEstado de registro\tFecha de registro\tAprobación (Polibid)"
    )
    for p in data["participantes"]:
        nombre = p.get("nombre_principal") or "—"
        correo = p.get("email") or "—"
        usuario = f"{nombre} · {correo}"
        lines.append(
            f"{p['id']}\t{usuario}\t{p.get('status')}\t{_fmt_dt(p.get('created_at'))}\t{_fmt_dt(p.get('approved_at'))}"
        )
    parts_ap = []
    for p in data["participantes"]:
        tag = (p.get("email") or "—").strip()
        parts_ap.append(f"{tag} {_fmt_dt(p.get('approved_at'))}")
    if parts_ap:
        lines.append("")
        lines.append("Aprobación en Polibid:\t" + " · ".join(parts_ap))
    lines.append("")
    lines.append("Historial de pujas (orden cronológico)")
    lines.append(
        "Usuario (nombre · correo)\tMonto (COP)\tTipo (agente)\tIP\tFecha y hora\tEstado"
    )
    hp = data.get("historial_pujas") or []
    if not hp:
        lines.append("—\t(Sin pujas registradas)")
    else:
        for row in hp:
            lines.append(
                f"{row['usuario']}\t{_format_cop_co(row.get('monto'))}\t{row.get('tipo')}\t"
                f"{row.get('ip')}\t{_fmt_dt(row.get('fecha'))}\t{row.get('estado')}"
            )
    lines.append("")
    lines.append("Ganador y oferta ganadora")
    g = data.get("ganador")
    lines.append("Campo\tValor")
    if g:
        gan = (
            f"{g.get('nombre_principal') or '—'} · correo {g.get('email') or '—'} · "
            f"id participante {g.get('participant_id')}"
        )
        oferta = (
            f"{_format_cop_co(g.get('amount'))} · id puja {g.get('bid_id')} · estado {g.get('status')} · "
            f"{g.get('created_at')}"
        )
        lines.append(f"Participante ganador\t{gan}")
        lines.append(f"Oferta ganadora\t{oferta}")
    else:
        lines.append("—\t(Sin ofertas registradas o no hay datos de ganador)")
    return "\n".join(lines)


def informe_teams_html(data: dict[str, Any]) -> str:
    """HTML con tablas claras (vista previa y base para Teams)."""
    r, s = data["resumen"], data["subasta"]
    # Estilos: bordes visibles, celdas separadas, UTF-8 en contenido escapado
    wrap = (
        'style="font-family:\'Segoe UI\',Roboto,Arial,sans-serif;font-size:14px;'
        'color:#242424;line-height:1.45;max-width:960px;"'
    )
    th = (
        'style="border:1px solid #605e5c;padding:10px 12px;text-align:left;'
        'background:#f3f2f1;font-weight:600;vertical-align:top;"'
    )
    td = 'style="border:1px solid #605e5c;padding:10px 12px;vertical-align:top;background:#ffffff;"'
    td_alt = 'style="border:1px solid #605e5c;padding:10px 12px;vertical-align:top;background:#faf9f8;"'
    td_lbl = (
        'style="border:1px solid #605e5c;padding:10px 12px;vertical-align:top;'
        'background:#edebe9;font-weight:600;width:28%;"'
    )
    h3 = 'style="margin:22px 0 10px 0;font-size:16px;color:#201f1e;border-bottom:2px solid #c8c6c4;padding-bottom:6px;"'
    tbl = 'style="border-collapse:collapse;width:100%;table-layout:fixed;margin-bottom:8px;"'

    sb = [
        f"<div {wrap}>",
        f"<p {h3}><strong>Informe de subasta</strong></p>",
        f"<p style=\"margin:0 0 16px 0;\"><strong>Identificador:</strong> {_esc(data['auction_id'])}</p>",
        f"<p {h3}>Datos de la subasta</p>",
        f"<table {tbl} role=\"presentation\">",
        f"<tr><td {td_lbl}>Código</td><td {td}>{_esc(s.get('code'))}</td></tr>",
        f"<tr><td {td_lbl}>Estado</td><td {td_alt}>{_esc(s.get('status'))}</td></tr>",
        f"<tr><td {td_lbl}>Descripción</td><td {td}>{_esc(s.get('title'))}</td></tr>",
        "</table>",
        f"<p {h3}>Resumen</p>",
        f"<table {tbl}><thead><tr>",
        f"<th {th}>Concepto</th><th {th}>Valor</th>",
        "</tr></thead><tbody>",
        f"<tr><td {td}>Participantes inscritos</td><td {td}>{_esc(r['participantes_inscritos'])}</td></tr>",
        "<tr><td "
        f"{td_alt}>Participantes con ofertas <span style=\"font-size:12px;color:#605e5c;\">(número de pujas)</span></td>"
        f"<td {td_alt}>{_esc(r['participantes_con_ofertas_numero_pujas'])}</td></tr>",
        f"<tr><td {td}>Total de ofertas recibidas</td><td {td}>{_esc(r['total_ofertas_recibidas'])}</td></tr>",
        "</tbody></table>",
        f"<p {h3}>Participantes</p>",
        f"<table {tbl}><thead><tr>",
        "<th style=\"border:1px solid #605e5c;padding:10px 12px;text-align:left;background:#f3f2f1;"
        'font-weight:600;vertical-align:top;width:22%;word-break:break-all;">ID participante</th>',
        "<th style=\"border:1px solid #605e5c;padding:10px 12px;text-align:left;background:#f3f2f1;"
        'font-weight:600;vertical-align:top;width:30%;\">Nombre y correo</th>',
        "<th style=\"border:1px solid #605e5c;padding:10px 12px;text-align:left;background:#f3f2f1;"
        'font-weight:600;vertical-align:top;width:12%;\">Estado</th>',
        "<th style=\"border:1px solid #605e5c;padding:10px 12px;text-align:left;background:#f3f2f1;"
        'font-weight:600;vertical-align:top;width:18%;\">Fecha de registro</th>',
        "<th style=\"border:1px solid #605e5c;padding:10px 12px;text-align:left;background:#f3f2f1;"
        'font-weight:600;vertical-align:top;width:18%;\">Fecha de aprobación</th>',
        "</tr></thead><tbody>",
    ]
    for i, p in enumerate(data["participantes"]):
        nombre = p.get("nombre_principal") or "—"
        correo = p.get("email") or "—"
        usuario = f"{nombre} · {correo}"
        rowtd = td if i % 2 == 0 else td_alt
        bg = "#ffffff" if i % 2 == 0 else "#faf9f8"
        id_cell = (
            f"style=\"border:1px solid #605e5c;padding:10px 12px;vertical-align:top;background:{bg};"
            'word-break:break-all;font-size:12px;"'
        )
        sb.append(
            "<tr>"
            f"<td {id_cell}>{_esc(p['id'])}</td>"
            f"<td {rowtd}>{_esc(usuario)}</td>"
            f"<td {rowtd}>{_esc(p.get('status'))}</td>"
            f"<td {rowtd}>{_esc(_fmt_dt(p.get('created_at')))}</td>"
            f"<td {rowtd}>{_esc(_fmt_dt(p.get('approved_at')))}</td>"
            "</tr>"
        )
    sb.append("</tbody></table>")
    parts_ap = []
    for p in data["participantes"]:
        tag = (p.get("email") or "—").strip()
        parts_ap.append(f"{tag} {_fmt_dt(p.get('approved_at'))}")
    if parts_ap:
        sb.append(
            f"<p style=\"margin:14px 0;padding:10px 12px;background:#fff4ce;border:1px solid #ffb900;border-radius:4px;\">"
            f"<strong>Aprobaciones:</strong> {_esc(' · '.join(parts_ap))}</p>"
        )
    sb.append(f"<p {h3}>Historial de pujas</p>")
    sb.append(
        "<p style=\"margin:0 0 10px 0;font-size:12px;color:#605e5c;\">"
        "Orden cronológico. <em>Tipo</em> es un recorte del agente de usuario (navegador) registrado en la puja."
        "</p>"
    )
    sb.append(f"<table {tbl}><thead><tr>")
    sb.append(
        f"<th {th}>Usuario</th>"
        f"<th {th}>Monto (COP)</th>"
        f"<th {th}>Tipo</th>"
        f"<th {th}>IP</th>"
        f"<th {th}>Fecha y hora</th>"
        f"<th {th}>Estado</th>"
        "</tr></thead><tbody>"
    )
    hp = data.get("historial_pujas") or []
    if not hp:
        sb.append(
            f"<tr><td {td} colspan=\"6\" style=\"text-align:center;color:#605e5c;\">"
            "(Sin pujas registradas)</td></tr>"
        )
    else:
        for i, row in enumerate(hp):
            rowtd = td if i % 2 == 0 else td_alt
            base_style = rowtd.rstrip('"')
            tipo_style = f'{base_style};font-size:12px;word-break:break-word;"'
            sb.append(
                "<tr>"
                f"<td {rowtd}>{_esc(row.get('usuario'))}</td>"
                f"<td {rowtd}>{_esc(_format_cop_co(row.get('monto')))}</td>"
                f"<td {tipo_style}>{_esc(row.get('tipo'))}</td>"
                f"<td {rowtd}>{_esc(row.get('ip'))}</td>"
                f"<td {rowtd}>{_esc(_fmt_dt(row.get('fecha')))}</td>"
                f"<td {rowtd}>{_esc(row.get('estado'))}</td>"
                "</tr>"
            )
    sb.append("</tbody></table>")
    sb.append(f"<p {h3}>Ganador y oferta ganadora</p>")
    sb.append(f"<table {tbl}><thead><tr><th {th}>Campo</th><th {th}>Detalle</th></tr></thead><tbody>")
    g = data.get("ganador")
    if g:
        gan = (
            f"{g.get('nombre_principal') or '—'} · correo {g.get('email') or '—'} · "
            f"id participante {g.get('participant_id')}"
        )
        oferta = (
            f"{_format_cop_co(g.get('amount'))} · id puja {g.get('bid_id')} · estado {g.get('status')} · "
            f"{g.get('created_at')}"
        )
        sb.append(
            f"<tr><td {td}>Participante ganador</td><td {td}>{_esc(gan)}</td></tr>"
            f"<tr><td {td_alt}>Oferta ganadora</td><td {td_alt}>{_esc(oferta)}</td></tr>"
        )
    else:
        sb.append(
            f"<tr><td {td}>—</td><td {td}>(Sin ofertas registradas o sin datos de ganador)</td></tr>"
        )
    sb.append("</tbody></table></div>")
    return "".join(sb)


def informe_teams_html_clipboard(data: dict[str, Any]) -> str:
    """Documento HTML completo UTF-8 para portapapeles (Teams / Word)."""
    inner = informe_teams_html(data)
    return (
        "<!DOCTYPE html><html><head>"
        '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
        "<meta charset=\"utf-8\">"
        "</head><body>"
        "<!--StartFragment-->"
        f"{inner}"
        "<!--EndFragment-->"
        "</body></html>"
    )


def format_informe_texto(data: dict[str, Any]) -> str:
    """Mismo contenido que el portapapeles texto de Teams (CLI)."""
    return informe_teams_plain(data)


def informe_teams_clipboard_payload(data: dict[str, Any]) -> tuple[str, str]:
    """(html completo UTF-8, texto plano) para ClipboardItem."""
    return informe_teams_html_clipboard(data), informe_teams_plain(data)


def informe_teams_clipboard_b64(data: dict[str, Any]) -> tuple[str, str]:
    """Base64 UTF-8 para incrustar en HTML/JS sin problemas de escape."""
    h, p = informe_teams_clipboard_payload(data)
    return base64.b64encode(h.encode("utf-8")).decode("ascii"), base64.b64encode(p.encode("utf-8")).decode(
        "ascii"
    )

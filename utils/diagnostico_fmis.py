import sys
from pathlib import Path
PROYECTO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROYECTO / "core"))
from core import get_connection, load_dotenv_files
from vault import read_vault, decrypt_payload, dsn_from_database_section
load_dotenv_files()

local_conf_path = PROYECTO / "local.conf"
vault_path = PROYECTO / "credentials.vault.enc"
dsn = None
if not local_conf_path.is_file() and vault_path.is_file():
    password = input("Contraseña del vault: ").strip()
    blob = read_vault(vault_path)
    creds = decrypt_payload(blob, password)
    db = creds.get("database") or creds
    dsn = dsn_from_database_section(db)

conn = get_connection(dsn, local_conf_path if local_conf_path.is_file() else None)
auction_uuid = "4cfcef04-f3e8-41c5-ace7-071d157b74b8"

with conn.cursor() as cur:
    # Buscar por manifestacion_interes de participantes
    cur.execute("""
        SELECT DISTINCT mani.grupo_id, mani.inmueble_id
        FROM polybid.auction_participants p
        JOIN polibid_credentials pc ON pc.client_id = p.client_id
        JOIN manifestacion_interes mani ON mani.contact_tercero_id = pc.contact_tercero_id
        WHERE p.auction_id = %s::uuid
        AND (mani.grupo_id IS NOT NULL OR mani.inmueble_id IS NOT NULL)
        ORDER BY mani.grupo_id DESC NULLS LAST
    """, (auction_uuid,))
    rows = cur.fetchall()
    print(f"Manifestaciones: {rows}")

    for grupo_id, inm_id in rows:
        if grupo_id:
            print(f"\nGrupo {grupo_id} — inmuebles:")
            cur.execute("""
                SELECT id, codigo, numero_matricula, nombre_grupo, codigo_grupo, es_padre
                FROM mst_inmuebles WHERE grupo_id = %s ORDER BY es_padre DESC, id
            """, (grupo_id,))
            for r in cur.fetchall():
                print(f"  id={r[0]} codigo={r[1]} fmi={r[2]} nombre={r[3]} padre={r[5]}")
        elif inm_id:
            cur.execute("SELECT id, codigo, numero_matricula, nombre_grupo, codigo_grupo FROM mst_inmuebles WHERE id=%s", (inm_id,))
            r = cur.fetchone()
            print(f"\nInmueble individual: {r}")

conn.close()

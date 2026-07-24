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
    password = input("Contraseña: ").strip()
    blob = read_vault(vault_path)
    creds = decrypt_payload(blob, password)
    db = creds.get("database") or creds
    dsn = dsn_from_database_section(db)
conn = get_connection(dsn, local_conf_path if local_conf_path.is_file() else None)

auction_uuid = "8cc8d7f3-118a-4658-b757-1567985a8620"  # subasta 46

with conn.cursor() as cur:
    # Buscar por manifestaciones
    cur.execute("""
        SELECT mani.grupo_id, mani.inmueble_id, COUNT(*) as cnt
        FROM polybid.auction_participants p
        JOIN polibid_credentials pc ON pc.client_id = p.client_id
        JOIN manifestacion_interes mani ON mani.contact_tercero_id = pc.contact_tercero_id
        WHERE p.auction_id = %s::uuid
        AND (mani.grupo_id IS NOT NULL OR mani.inmueble_id IS NOT NULL)
        GROUP BY mani.grupo_id, mani.inmueble_id
        ORDER BY cnt DESC
    """, (auction_uuid,))
    print("Manifestaciones:")
    for r in cur.fetchall():
        print(f"  grupo_id={r[0]} inmueble_id={r[1]} count={r[2]}")

    # Ver inmueble que esta agarrando
    cur.execute("""
        SELECT id, codigo, numero_matricula, grupo_id, codigo_grupo, es_padre
        FROM mst_inmuebles WHERE grupo_id = 442 ORDER BY es_padre DESC, id
    """)
    print("\nInmuebles grupo 442:")
    for r in cur.fetchall():
        print(f"  id={r[0]} codigo={r[1]} fmi={r[2]} grupo_id={r[3]} codigo_grupo={r[4]} padre={r[5]}")

conn.close()

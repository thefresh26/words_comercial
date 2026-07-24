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

busqueda = input("Cédula o nombre: ").strip()

with conn.cursor() as cur:
    if busqueda.isdigit():
        cur.execute("""
            SELECT ct.nombre_principal, ct.identificacion_numero, ct.identificacion_tipo,
                   ct.lugar_expedicion_doc, ct.fecha_diligenciamiento
            FROM contact_terceros ct
            WHERE ct.identificacion_numero = %s
        """, (busqueda,))
    else:
        cur.execute("""
            SELECT ct.nombre_principal, ct.identificacion_numero, ct.identificacion_tipo,
                   ct.lugar_expedicion_doc, ct.fecha_diligenciamiento
            FROM contact_terceros ct
            WHERE ct.nombre_principal ILIKE %s
            LIMIT 10
        """, (f"%{busqueda}%",))
    
    rows = cur.fetchall()
    if not rows:
        print("No encontrado en contact_terceros")
    else:
        print(f"\nEncontrado(s): {len(rows)}")
        for r in rows:
            print(f"  Nombre:  {r[0]}")
            print(f"  Cédula:  {r[1]} ({r[2]})")
            print(f"  Ciudad:  {r[3]}")
            print(f"  Fecha:   {r[4]}")
            print()

        # Buscar subastas
        cur.execute("""
            SELECT a.code, a.id, a.status
            FROM contact_terceros ct
            JOIN polibid_credentials pc ON pc.contact_tercero_id = ct.id
            JOIN polybid.auction_participants p ON p.client_id = pc.client_id
            JOIN polybid.auctions a ON a.id = p.auction_id
            WHERE ct.identificacion_numero = %s OR ct.nombre_principal ILIKE %s
            ORDER BY a.created_at DESC
        """, (busqueda, f"%{busqueda}%"))
        subastas = cur.fetchall()
        if subastas:
            print("Subastas en las que participó:")
            for s in subastas:
                print(f"  {s[0]} — {s[2]}")
        else:
            print("No ha participado en ninguna subasta registrada.")

conn.close()

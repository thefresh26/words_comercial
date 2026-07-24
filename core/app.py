"""
Interfaz web: primero copiar para Teams, luego vista previa del informe.

  python run.py
  ./run.sh
"""

from __future__ import annotations

import sys
from pathlib import Path

_pkg = Path(__file__).resolve().parent
if str(_pkg) not in sys.path:
    sys.path.insert(0, str(_pkg))

import streamlit as st
import streamlit.components.v1 as components

from core import (
    fetch_informe,
    get_connection,
    informe_teams_clipboard_b64,
    informe_teams_html,
    load_dotenv_files,
    resolve_local_conf_path,
    vista_previa_extra_html,
)

st.set_page_config(page_title="Informe de subasta", layout="wide")

load_dotenv_files()
_LOCAL_CONF = resolve_local_conf_path(None)

st.title("Informe de subasta")
st.markdown(
    "Introduce el **identificador de la subasta** y pulsa **Consultar**. "
    "Lo primero podrás **copiar el informe para Teams**; más abajo está la vista previa."
)

with st.sidebar:
    st.markdown(
        """
**Cómo usar**
1. Pega el identificador (el que te dieron para esa subasta).
2. Pulsa **Consultar**.
3. **Copiar para Teams** y pega en el chat (Ctrl+V o ⌘V).
4. Si quieres, revisa la vista previa más abajo.
        """
    )

with st.form("consulta", clear_on_submit=False):
    st.subheader("Consulta")
    auction_id = st.text_input(
        "Identificador de la subasta",
        label_visibility="collapsed",
        placeholder="Pega aquí el identificador…",
        help="Es el código largo con guiones, como el que usas al generar el informe por otros medios.",
    )
    submitted = st.form_submit_button("Consultar", type="primary", use_container_width=False)

if submitted:
    aid = (auction_id or "").strip()
    if not aid:
        st.warning("Escribe o pega el identificador de la subasta.")
    else:
        try:
            conn = get_connection(None, _LOCAL_CONF)
        except ImportError as e:
            st.error(str(e))
        except Exception:
            st.error(
                "No se pudo obtener la información. Comprueba tu conexión a internet "
                "y la red de la empresa, o pide ayuda a quien administra el sistema."
            )
        else:
            try:
                data = fetch_informe(conn, aid)
            except ValueError:
                st.error("No encontramos una subasta con ese identificador.")
                data = None
            finally:
                conn.close()

            if data:
                st.session_state["informe_data"] = data

if "informe_data" in st.session_state:
    data = st.session_state["informe_data"]
    html_preview = informe_teams_html(data)
    h64, p64 = informe_teams_clipboard_b64(data)

    st.divider()
    st.subheader("Copiar para Teams")
    st.caption("El botón guarda tablas y textos con acentos y eñes correctamente al pegar en Teams.")

    # Decodificación UTF-8 real del base64 (atob solo sirve para ASCII)
    components.html(
        f"""
<div style="font-family:Segoe UI,Arial,sans-serif;padding:4px 0;">
  <button id="cpyTeams" type="button"
    style="padding:12px 24px;font-size:16px;font-weight:600;cursor:pointer;
    background:#6264A7;color:#fff;border:none;border-radius:6px;">
    Copiar informe para Teams
  </button>
</div>
<script>
  function b64ToUtf8(b64) {{
    const raw = atob(b64);
    const bytes = new Uint8Array(raw.length);
    for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
    return new TextDecoder("utf-8").decode(bytes);
  }}
  const h64 = "{h64}";
  const p64 = "{p64}";
  const btn = document.getElementById("cpyTeams");
  btn.addEventListener("click", async function () {{
    const html = b64ToUtf8(h64);
    const plain = b64ToUtf8(p64);
    try {{
      await navigator.clipboard.write([
        new ClipboardItem({{
          "text/html": new Blob([html], {{ type: "text/html;charset=utf-8" }}),
          "text/plain": new Blob([plain], {{ type: "text/plain;charset=utf-8" }}),
        }}),
      ]);
      btn.textContent = "Copiado — abre Teams y pega (Ctrl+V)";
      btn.style.background = "#107C41";
    }} catch (err) {{
      try {{
        await navigator.clipboard.writeText(plain);
        btn.textContent = "Copiado como texto — pega en Teams";
        btn.style.background = "#CA5010";
      }} catch (e2) {{
        btn.textContent = "No se pudo copiar automáticamente";
      }}
    }}
  }});
</script>
""",
        height=140,
    )

    st.divider()
    st.subheader("Vista previa")
    st.markdown(html_preview, unsafe_allow_html=True)
    st.markdown(
        vista_previa_extra_html(data.get("vista_previa_extra") or {}),
        unsafe_allow_html=True,
    )

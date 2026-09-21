# Automatización de Documentos — Comercial

Herramientas para generar automáticamente los documentos de comercial (actas de subasta, certificados de debida diligencia, informes y declaraciones juramentadas) a partir de la base de datos de Polybid, y una interfaz web para consultar informes y copiarlos a Microsoft Teams.

## Estructura del proyecto

```
core/                 Lógica compartida y app web
  core.py               Conexión a BD (local.conf / .env) y consulta del informe de subasta
  vault.py              Cifrado/descifrado de credenciales (credentials.vault.enc)
  app.py                Interfaz web (Streamlit)

scripts/               Generadores de documentos (uno por tipo de documento)
  generar_acta.py
  generar_certificados_dd.py
  generar_informe.py
  generar_juramentadas.py

utils/               Herramientas auxiliares
  informe_subasta_polybid.py   (CLI del informe, sin interfaz web)
  clarity_capture.py            (estadísticas de tráfico de Microsoft Clarity)

templates/              Plantillas .docx originales usadas para generar cada documento
  002_CERTIFICADO_RESULTADO_DD.docx
  003_FORMATO_DECLARACION_JURAMENTADA_FO_GP_008.docx
  ACTA_DE_CERTIFICACIÓN_DE_SUBASTA_ELECTRÓNICA.docx
  INFORME_SUBASTA.docx

output/                 Documentos generados (no se versiona en git)
  actas/
  certificados_dd/
  informes/
  juramentadas/

run.py, run.sh, run.bat    Arranque de la interfaz web (venv + dependencias + credenciales)
requirements.txt           Dependencias Python
credentials.vault.enc       Credenciales de BD cifradas (no se versiona en git)
local.conf                  Credenciales de BD en claro para desarrollo (no se versiona en git)
```

## Requisitos

- Python 3.10 o superior
- Acceso a la base de datos de Polybid (VPN de la empresa si aplica)

## Configuración de credenciales de base de datos

El proyecto soporta dos formas de conectarse a la base de datos (se busca en este orden):

1. **`credentials.vault.enc`** en la raíz del proyecto — credenciales cifradas con contraseña (`core/vault.py`). Es lo que usan los operadores; la contraseña se solicita al arrancar `run.py`.
2. **`local.conf`** en la raíz del proyecto — JSON en claro con una sección `"database"` (`user`, `password`, `address`, `database`, `parameters`). Pensado para desarrollo local.

También se puede definir `DATABASE_URL` (o variables `PG*`) en un archivo `.env` en la raíz.

Ninguno de estos archivos se versiona en git (ver `.gitignore`).

## Uso

### Interfaz web (recomendado para operadores)

```bash
python run.py
# o en Windows: run.bat
# o en Mac/Linux: ./run.sh
```

Esto crea/activa un entorno virtual (`.venv`), instala dependencias, pide la contraseña del vault (si existe) y abre `http://localhost:8501` con la interfaz para consultar el informe de una subasta y copiarlo a Teams.

### Generar documentos por línea de comandos

Cada script recibe el UUID de la subasta, su código (ACTIBID-...), un
FMI/número de matrícula, el código de un inmueble individual, o el código
de una unidad inmobiliaria (UNI-XXXX-AAAA). Si el FMI/unidad tiene varias
subastas asociadas, el script las lista y pide elegir cuál usar:

```bash
python scripts/generar_acta.py <identificador>
python scripts/generar_certificados_dd.py <identificador>
python scripts/generar_informe.py <identificador>
python scripts/generar_juramentadas.py <identificador|cedula>
```

Los documentos generados se guardan en `output/actas/`, `output/certificados_dd/`, `output/informes/` y `output/juramentadas/` respectivamente.

### Informe por CLI (sin interfaz web)

```bash
python utils/informe_subasta_polybid.py <auction_uuid> --json
```

### Instalación manual de dependencias (opcional)

```bash
pip install -r requirements.txt
```

## Notas

- `output/` no se versiona: los documentos generados quedan solo en tu máquina.
- `credentials.vault.enc` y `local.conf` tampoco se versionan por contener (o dar acceso a) credenciales de base de datos.

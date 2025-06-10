# 🤖 Automatización de Reportes para Slack

Este es un proyecto personal construido con el objetivo de automatizar un flujo de trabajo completo, desde la lectura de datos en archivos PDF hasta la distribución de reportes personalizados a través de la **API de Slack**. El sistema extrae transacciones de consumo, genera reportes individuales y los distribuye de forma privada a cada vendedor usando un bot.

---

## 🎯 Objetivo del Proyecto

Este proyecto fue creado como una solución práctica para:

- Automatizar una tarea manual y repetitiva de extracción y envío de información.
- Aprender a interactuar con archivos complejos como PDF y generar reportes en formatos como imagen y Excel.
- Construir una integración robusta con la **API de Slack**, manejando autenticación, permisos y envío de mensajes y archivos.
- Practicar la creación de scripts de Python modulares y configurables para un caso de uso real.

---

## 🛠️ Tecnologías Utilizadas

### Lenguaje Principal
- 🐍 **Python**

### Librerías Clave
- 💬 **slack-sdk**: Para toda la comunicación con la API de Slack.
- 📄 **PyMuPDF**: Para la extracción de datos de alto rendimiento desde archivos PDF.
- 🎨 **Pillow**: Para la creación y manipulación de las imágenes de los reportes.
- 🐼 **Pandas** & **openpyxl**: Para la generación de reportes consolidados en formato Excel.
- 🌐 **requests**: Para realizar peticiones HTTP necesarias en el flujo del bot.

---

## 📁 Estructura del Proyecto

```bash
automatizacion-reportes-slack/
├── reportes_img/            # Carpeta donde se guardan las imágenes generadas
├── env/                     # Entorno virtual de Python
├── pdf_extractor_excel.py   # Script que genera el reporte en Excel/TXT
├── pdf_extractor_imagenes.py# Script que genera las imágenes por vendedor
├── bot_lector.py            # Script principal del bot de Slack (modo polling)
├── vendedores.json          # Archivo de configuración de vendedores y UIDs
├── consumos.txt             # Archivo de texto generado con los datos
├── requirements.txt         # Lista de dependencias del proyecto
└── README.md                # Este archivo
```

---

## ⚙️ Configuración e Instalación
### Prerrequisitos
- Python 3.8 o superior instalado.
- Una cuenta y un workspace de Slack con permisos de administrador para poder crear y configurar un bot.

### Pasos
1. Clona el repositorio:
```bash
git clone [https://github.com/tu-usuario/tu-repositorio.git](https://github.com/tu-usuario/tu-repositorio.git)
cd tu-repositorio
```

2. Crea y activa un entorno virtual:
```Bash

# Crear el entorno
python -m venv env

# Activar en Windows
.\env\Scripts\activate

# Activar en macOS/Linux
source env/bin/activate
```

3. Instala las dependencias:
Crea un archivo requirements.txt con el contenido de abajo y luego ejecuta pip.

```Plaintext
requests
PyMuPDF
Pillow
pandas
openpyxl
slack_sdk
```

```Bash
pip install -r requirements.txt
```

#### 4. Configura tu Bot y Archivos:

**A. Crear el Bot en Slack**

1.  Ve a `https://api.slack.com/apps` y crea una nueva aplicación.
2.  En "OAuth & Permissions" -> "Scopes", añade los siguientes permisos de Bot:
    -   `groups:history` (para leer mensajes en canales privados)
    -   `files:read` (para leer información de archivos)
    -   `files:write` (para subir/reenviar archivos)
    -   `chat:write` (para enviar mensajes)
    -   `im:write` (para iniciar mensajes directos)
3.  Instala la aplicación en tu workspace y copia el "Bot User OAuth Token" (empieza con `xoxb-`).
4.  Invita a tu bot al canal privado que usarás para subir los archivos (`@[Nombre del bot]`).

**B. Configurar las Variables de Entorno (Método Seguro para el Token)**

Para evitar exponer tu token, no lo escribiremos en el código. En su lugar, lo configuraremos como una variable de entorno en tu terminal cada vez que vayas a ejecutar el script.

En Windows (PowerShell):
```PowerShell
$env:SLACK_BOT_TOKEN = "PEGA-AQUI-EL-NUEVO-TOKEN-QUE-ACABAS-DE-COPIAR"
```

En macOS/Linux:
```Bash
export SLACK_BOT_TOKEN="PEGA-AQUI-EL-NUEVO-TOKEN-QUE-ACABAS-DE-COPIAR"
```
(El script ya está preparado para leer esta variable. No necesitas modificar el código Python para esto)

**C. Configurar Archivos Locales**
JSON de Vendedores: Crea y rellena el archivo vendedores.json con los nombres y los IDs de usuario (UID) de Slack de cada vendedor.

---

## ▶️ Modo de Uso
1. Genera los reportes: Ejecuta los scripts pdf_extractor_... para crear los archivos de imagen y Excel a partir de los PDFs originales.
2. Sube las imágenes a Slack: Arrastra y suelta las imágenes generadas (Consumos VENDEDOR.jpg) en el canal privado que has configurado para el bot.
3. Activa el bot: En tu terminal (con el entorno virtual activado), ejecuta el script principal y déjalo corriendo.
```Bash
python bot_lector.py
```
4. El bot detectará los archivos nuevos, los procesará y los enviará a cada vendedor. Cuando termines, puedes detener el script con Ctrl+C

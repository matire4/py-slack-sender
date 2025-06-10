import json
import time
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

# --- CONFIGURACIÓN ---
SLACK_BOT_TOKEN = os.getenv("SLACK_BOT_TOKEN") 
JSON_FILE_PATH = "user_id.json"
CONSUMOS_TXT_PATH = "output_txt/consumos.txt"
# --- FIN DE LA CONFIGURACIÓN ---


def parsear_consumos(ruta_archivo_txt):
    """
    Lee un archivo de texto con múltiples reportes y los separa en un
    diccionario, donde la clave es el nombre del vendedor.
    """
    print(f"📄 Leyendo y procesando el archivo '{ruta_archivo_txt}'...")
    try:
        with open(ruta_archivo_txt, 'r', encoding='utf-8') as f:
            contenido_completo = f.read()
    except FileNotFoundError:
        print(f"🔴 ERROR: No se encontró el archivo de texto: '{ruta_archivo_txt}'")
        return None

    reportes = {}
    bloques = contenido_completo.split('--- Consumos')[1:]

    for bloque in bloques:
        try:
            lineas = bloque.strip().split('\n')
            nombre_vendedor = lineas[0].replace('---', '').strip()
            cuerpo_reporte = '\n'.join(lineas[1:]).strip()
            reportes[nombre_vendedor] = cuerpo_reporte
            print(f"   - Reporte de '{nombre_vendedor}' extraído.")
        except IndexError:
            print("   - 🟡 Se encontró un bloque de texto malformado. Omitiendo.")
            continue
            
    return reportes


def enviar_reportes_de_texto():
    """
    Función principal que orquesta la lectura del TXT y el envío a Slack.
    """
    reportes_por_vendedor = parsear_consumos(CONSUMOS_TXT_PATH)
    if not reportes_por_vendedor:
        print("❌ No se encontraron reportes para procesar. El script ha terminado.")
        return

    try:
        with open(JSON_FILE_PATH, 'r', encoding='utf-8') as f:
            data_vendedores = json.load(f)
        print(f"✅ Datos de vendedores cargados desde '{JSON_FILE_PATH}'.")
    except FileNotFoundError:
        print(f"🔴 ERROR CRÍTICO: No se encontró el archivo JSON: '{JSON_FILE_PATH}'.")
        return

    client = WebClient(token=SLACK_BOT_TOKEN)
    print("🤖 Conectando a Slack...")

    print("\n--- Empezando a enviar reportes por Slack ---")
    for vendor_name_from_txt, reporte_texto in reportes_por_vendedor.items():
        
        # --- INICIO DE LA MODIFICACIÓN ---
        # Búsqueda inteligente que ignora mayúsculas/minúsculas y espacios
        vendedor_encontrado_info = None
        for json_key, json_info in data_vendedores.items():
            if json_key.strip().upper() == vendor_name_from_txt.strip().upper():
                vendedor_encontrado_info = json_info
                break # Detenemos la búsqueda cuando encontramos la primera coincidencia
        # --- FIN DE LA MODIFICACIÓN ---

        if not (vendedor_encontrado_info and vendedor_encontrado_info.get("send_message")):
            print(f"🟡 Omitiendo a '{vendor_name_from_txt}' (no encontrado en JSON o envío desactivado).")
            continue

        user_id = vendedor_encontrado_info.get("UID")
        if not user_id:
            print(f"🔴 Omitiendo a '{vendor_name_from_txt}' (no tiene UID en el JSON).")
            continue

        mensaje_formateado = (
            f"¡Hola {vendor_name_from_txt.title()}! 👋 Aquí tienes tu resumen de consumos de este mes:\n\n"
            f"```\n"
            f"{reporte_texto}\n"
            f"```"
        )

        try:
            print(f"  Enviando reporte a {vendor_name_from_txt} (ID: {user_id})...")
            client.chat_postMessage(channel=user_id, text=mensaje_formateado)
            print(f"  ✅ ¡Éxito! Reporte enviado a {vendor_name_from_txt}.")
        except SlackApiError as e:
            print(f"  🔴 ¡ERROR al enviar a {vendor_name_from_txt}! Causa: {e.response['error']}")
        
        time.sleep(1)

    print("\n✅ Proceso completado. Todos los reportes han sido procesados.")

if __name__ == "__main__":
    enviar_reportes_de_texto()
"""
WaitSender - Powered by Teinek Solutions
==============================================================================
Suite de Automatización con Control Estricto de Envíos y Pendientes:
- Generación garantizada de Excel de reintentos con TODOS los contactos no completados.
- Despacho sin recargas mediante atajo nativo 'Nuevo Chat' (Ctrl + Alt + N).
- Compatibilidad con registrados y no registrados.
- Detección de movimiento de mouse (Pausa/Reanudación).
- Diseño Flat Minimalista sin marcos duros.
==============================================================================
"""

from __future__ import annotations

# ==============================================================================
# 1. IMPORTACIONES
# ==============================================================================
import base64
from datetime import datetime, timedelta
import io
import json
import os
import re
import shutil
import subprocess
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional
import urllib.parse
import webbrowser

from apscheduler.schedulers.background import BackgroundScheduler
import pandas as pd
import requests

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    Image, ImageTk = None, None
    PIL_AVAILABLE = False

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_FILES, TkinterDnD = None, None
    DND_AVAILABLE = False

try:
    import pyautogui
except ImportError:
    pyautogui = None

try:
    import speech_recognition as sr
except ImportError:
    sr = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    import docx
except ImportError:
    docx = None

try:
    from groq import Groq
except ImportError:
    Groq = None

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

try:
    import anthropic
except ImportError:
    anthropic = None

# ==============================================================================
# 2. RUTAS Y CONSTANTES DEL SISTEMA
# ==============================================================================
APP_DIR: str = os.path.abspath(os.path.dirname(__file__))
DIR_HISTORIAL: str = os.path.join(APP_DIR, "historial_mensajes")
DIR_FALLOS: str = os.path.join(APP_DIR, "reporte_fallos")
DIR_CONOCIMIENTO: str = os.path.join(APP_DIR, "base_conocimientos_ia")
DIR_RECURSOS: str = os.path.join(APP_DIR, "imagenes_campana")
DIR_CONTACTOS_VCF: str = os.path.join(APP_DIR, "contactos_exportados")

FILE_CONFIG: str = os.path.join(APP_DIR, "config_api.json")
FILE_BASES_DATOS: str = os.path.join(APP_DIR, "bases_de_datos.json")
FILE_HISTORIAL: str = os.path.join(DIR_HISTORIAL, "historial_general.json")
FILE_ESTRELLA: str = os.path.join(APP_DIR, "estrella.png")

TIEMPO_MINIMO_BLOQUE: int = 40

for path in (DIR_HISTORIAL, DIR_FALLOS, DIR_CONOCIMIENTO, DIR_RECURSOS, DIR_CONTACTOS_VCF):
    os.makedirs(path, exist_ok=True)

MODELOS_IA: Dict[str, str] = {
    "Groq": "openai/gpt-oss-120b",
    "OpenAI": "gpt-4o-mini",
    "Anthropic (Claude)": "claude-3-5-sonnet-20241022",
    "Demo / Offline": "Simulador Local",
}

FONT_FAMILY: str = "Segoe UI"

# ==============================================================================
# 3. PALETAS DE COLOR ULTRA-MINIMALISTAS (100% FLAT)
# ==============================================================================
THEMES: Dict[str, Dict[str, str]] = {
    "dark": {
        "bg_main": "#0B0E14",
        "bg_card": "#131A26",
        "bg_input": "#1C2638",
        "text_main": "#F8FAFC",
        "text_muted": "#94A3B8",
        "primary": "#1A4B8C",
        "primary_hover": "#163C70",
        "accent": "#38BDF8",
        "btn_now": "#2563EB",
        "btn_now_hover": "#1D4ED8",
        "danger": "#EF4444",
        "font_family": FONT_FAMILY,
    },
    "light": {
        "bg_main": "#F8FAFC",
        "bg_card": "#FFFFFF",
        "bg_input": "#F1F5F9",
        "text_main": "#0F172A",
        "text_muted": "#64748B",
        "primary": "#1A4B8C",
        "primary_hover": "#2563EB",
        "accent": "#0284C7",
        "btn_now": "#0284C7",
        "btn_now_hover": "#0369A1",
        "danger": "#DC2626",
        "font_family": FONT_FAMILY,
    },
    "blue_filter": {
        "bg_main": "#231F1A",
        "bg_card": "#2E2822",
        "bg_input": "#3D352E",
        "text_main": "#FDF6E2",
        "text_muted": "#C2B4A3",
        "primary": "#8C5E1A",
        "primary_hover": "#6E4812",
        "accent": "#D8A863",
        "btn_now": "#A27027",
        "btn_now_hover": "#7D541B",
        "danger": "#C2410C",
        "font_family": FONT_FAMILY,
    },
}

# ==============================================================================
# 4. PERSISTENCIA Y REPORTE ESTRICTO DE PENDIENTES
# ==============================================================================
def validar_y_limpiar_telefono(numero: str) -> Optional[str]:
    limpio = re.sub(r"[^\d]", "", str(numero))
    if 10 <= len(limpio) <= 15:
        return limpio
    return None

def cargar_configuracion() -> Dict[str, Any]:
    if os.path.exists(FILE_CONFIG):
        try:
            with open(FILE_CONFIG, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {
        "green_id": "",
        "green_token": "",
        "canal_predeterminado": "WhatsApp Web (Gratis)",
        "timeout_wa_web": 18,
        "llm_provider": "Groq",
        "llm_key": "",
        "llm_model": MODELOS_IA["Groq"],
        "perfil_nombre": "Nombre",
        "perfil_rol": "Especialista en Automatización",
        "perfil_empresa": "Teinek Solutions",
        "perfil_bibliografia": "Estilo profesional, claro y enfocado en soluciones ágiles.",
        "ultimos_manuales": [],
        "tema_actual": "dark",
        "categorias_bd": ["General", "Clientes", "Prospectos", "Cobranza"],
    }

def guardar_configuracion(config: Dict[str, Any]) -> None:
    with open(FILE_CONFIG, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)

def cargar_repositorio_bases() -> Dict[str, List[Dict[str, Any]]]:
    if os.path.exists(FILE_BASES_DATOS):
        try:
            with open(FILE_BASES_DATOS, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {"Base Principal": []}

def guardar_repositorio_bases(repo: Dict[str, List[Dict[str, Any]]]) -> None:
    with open(FILE_BASES_DATOS, "w", encoding="utf-8") as f:
        json.dump(repo, f, indent=4, ensure_ascii=False)

def exportar_contactos_a_vcf(contactos: List[Dict[str, Any]], nombre_archivo: str = "contactos_whatsapp.vcf") -> str:
    ruta_vcf = os.path.join(DIR_CONTACTOS_VCF, nombre_archivo)
    with open(ruta_vcf, "w", encoding="utf-8") as f:
        for c in contactos:
            nom = str(c.get("nombre", "Contacto")).strip() or "Contacto"
            num = str(c.get("numero", "")).strip()
            if num:
                f.write("BEGIN:VCARD\n")
                f.write("VERSION:3.0\n")
                f.write(f"FN:{nom}\n")
                f.write(f"TEL;TYPE=CELL:+{num}\n")
                f.write("END:VCARD\n")
    return ruta_vcf

def generar_reportes_dobles_fallos(fallidos: List[Dict[str, str]]) -> tuple[str, str]:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    archivo_txt = os.path.join(DIR_FALLOS, f"explicacion_fallos_{timestamp}.txt")
    with open(archivo_txt, "w", encoding="utf-8") as f:
        f.write("=" * 70 + "\n")
        f.write("REPORTE DE INCIDENCIAS Y PENDIENTES - WAITSENDER\n")
        f.write(f"Fecha y Hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Total de contactos pendientes de envío: {len(fallidos)}\n")
        f.write("=" * 70 + "\n\n")
        f.write("DETALLE POR CONTACTO:\n")
        f.write("-" * 70 + "\n")
        for i, item in enumerate(fallidos, 1):
            f.write(f"{i}. NOMBRE : {item.get('Nombre', 'Sin nombre')}\n")
            f.write(f"   NÚMERO : {item.get('Numero', 'Desconocido')}\n")
            f.write(f"   MOTIVO : {item.get('Motivo', 'Pendiente')}\n")
            f.write("-" * 50 + "\n")

    archivo_xlsx = os.path.join(DIR_FALLOS, f"reintentos_pendientes_{timestamp}.xlsx")
    datos_para_excel = []
    for item in fallidos:
        num = item.get("Numero", "").strip()
        if num and num != "N/A":
            datos_para_excel.append({
                "nombre": item.get("Nombre", "Amigo/a"),
                "telefono": num
            })

    if datos_para_excel:
        df_pendientes = pd.DataFrame(datos_para_excel)
        df_pendientes.to_excel(archivo_xlsx, index=False)
    else:
        pd.DataFrame([{"nombre": "Sin datos", "telefono": ""}]).to_excel(archivo_xlsx, index=False)

    return archivo_txt, archivo_xlsx

def respaldar_mensaje(destinatario: str, mensaje: str, canal: str) -> None:
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sanitized_name = "".join(c for c in destinatario if c.isalnum() or c in (" ", "_")).strip()
        filename = f"{timestamp}_{sanitized_name[:25]}.txt"
        filepath = os.path.join(DIR_HISTORIAL, filename)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"Canal: {canal}\nDestinatario: {destinatario}\nFecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n{mensaje}")
    except OSError:
        pass

def registrar_historial(entry: Dict[str, Any]) -> None:
    records = []
    if os.path.exists(FILE_HISTORIAL):
        try:
            with open(FILE_HISTORIAL, "r", encoding="utf-8") as f:
                records = json.load(f)
        except (json.JSONDecodeError, OSError):
            records = []
    records.append(entry)
    with open(FILE_HISTORIAL, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=4, ensure_ascii=False)

def obtener_historial() -> List[Dict[str, Any]]:
    if os.path.exists(FILE_HISTORIAL):
        try:
            with open(FILE_HISTORIAL, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return []
    return []

def copiar_imagen_powershell_seguro(path_imagen: str) -> bool:
    abs_path = os.path.abspath(path_imagen)
    if not os.path.exists(abs_path):
        return False
    ruta_escapada = abs_path.replace("'", "''")
    script = (
        "Add-Type -AssemblyName System.Windows.Forms; "
        "Add-Type -AssemblyName System.Drawing; "
        f"$img = [System.Drawing.Image]::FromFile([System.IO.Path]::GetFullPath('{ruta_escapada}')); "
        "[System.Windows.Forms.Clipboard]::SetImage($img); "
        "$img.Dispose();"
    )
    encoded_bytes = script.encode("utf-16le")
    encoded_str = base64.b64encode(encoded_bytes).decode("ascii")
    cmd = ["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded_str]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        return proc.returncode == 0
    except OSError:
        return False

# ==============================================================================
# 5. DISPATCHERS: SIN RECARGAS Y CON DETECCIÓN DE MOUSE
# ==============================================================================
def esperar_con_deteccion_mouse(segundos: int, callback_progreso: Optional[Any], evento_cancelar: Optional[threading.Event]) -> bool:
    tiempo_transcurrido = 0.0
    while tiempo_transcurrido < segundos:
        if evento_cancelar and evento_cancelar.is_set():
            return False
        
        pos_inicial = pyautogui.position() if pyautogui else (0, 0)
        time.sleep(0.5)
        pos_actual = pyautogui.position() if pyautogui else (0, 0)

        if abs(pos_actual[0] - pos_inicial[0]) > 6 or abs(pos_actual[1] - pos_inicial[1]) > 6:
            if callback_progreso:
                callback_progreso(-1, -1, "⏸️ Ratón en movimiento: Esperando reposo...")
            
            mouse_quieto = 0.0
            ultima_pos = pyautogui.position() if pyautogui else (0, 0)
            while mouse_quieto < 2.0:
                if evento_cancelar and evento_cancelar.is_set():
                    return False
                time.sleep(0.4)
                nueva_pos = pyautogui.position() if pyautogui else (0, 0)
                if abs(nueva_pos[0] - ultima_pos[0]) <= 3 and abs(nueva_pos[1] - ultima_pos[1]) <= 3:
                    mouse_quieto += 0.4
                else:
                    mouse_quieto = 0.0
                ultima_pos = nueva_pos

        tiempo_transcurrido += 0.5
    return True

def dispatch_whatsapp_web_sin_recargas(
    numero: str,
    mensaje: str,
    imagenes: Optional[List[str]] = None,
    timeout_espera: int = 18,
    es_primer_envio: bool = False,
    callback_progreso: Optional[Any] = None,
    evento_cancelar: Optional[threading.Event] = None,
) -> bool:
    if pyautogui is None:
        return False
    try:
        clean_num = re.sub(r"[^\d]", "", numero)

        if es_primer_envio:
            url = f"https://web.whatsapp.com/send?phone={clean_num}&text={urllib.parse.quote(mensaje)}"
            webbrowser.open(url)
            if not esperar_con_deteccion_mouse(max(15, timeout_espera), callback_progreso, evento_cancelar):
                return False
            
            if evento_cancelar and evento_cancelar.is_set():
                return False

            pyautogui.press("enter")
            time.sleep(2.5)
        else:
            if evento_cancelar and evento_cancelar.is_set():
                return False
            pyautogui.hotkey("ctrl", "alt", "n")
            time.sleep(1.0)

            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("backspace")
            time.sleep(0.3)

            subprocess.run(["powershell", "-command", f"Set-Clipboard -Value '{clean_num}'"], check=False)
            time.sleep(0.2)
            pyautogui.hotkey("ctrl", "v")
            
            if not esperar_con_deteccion_mouse(3, callback_progreso, evento_cancelar):
                return False

            pyautogui.press("enter")
            time.sleep(2.0)

            if evento_cancelar and evento_cancelar.is_set():
                return False

            subprocess.run(["powershell", "-command", f"Set-Clipboard -Value '{mensaje}'"], check=False)
            time.sleep(0.4)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.5)

            if evento_cancelar and evento_cancelar.is_set():
                return False

            pyautogui.press("enter")
            time.sleep(2.0)

        if imagenes:
            for img in imagenes:
                if evento_cancelar and evento_cancelar.is_set():
                    return False
                if os.path.exists(img) and copiar_imagen_powershell_seguro(img):
                    time.sleep(1.2)
                    pyautogui.hotkey("ctrl", "v")
                    time.sleep(3.0)
                    pyautogui.press("enter")
                    time.sleep(2.0)

        if evento_cancelar and evento_cancelar.is_set():
            return False

        return True
    except Exception as e:
        print(f"Error en despacho de {numero}: {e}")
        return False

def dispatch_green_api(instancia: str, token: str, numero: str, mensaje: str, imagenes: Optional[List[str]] = None) -> bool:
    if not instancia or not token:
        return False
    clean_num = re.sub(r"[^\d]", "", numero)
    try:
        url_text = f"https://api.green-api.com/waInstance{instancia}/sendMessage/{token}"
        payload_text = {"chatId": f"{clean_num}@c.us", "message": mensaje}
        res_text = requests.post(url_text, json=payload_text, timeout=15)
        success = res_text.status_code == 200

        if imagenes:
            url_file = f"https://api.green-api.com/waInstance{instancia}/sendFileByUpload/{token}"
            for img_path in imagenes:
                if os.path.exists(img_path):
                    filename = os.path.basename(img_path)
                    with open(img_path, "rb") as f_img:
                        files = {"file": (filename, f_img, "image/jpeg")}
                        payload = {"chatId": f"{clean_num}@c.us", "fileName": filename}
                        requests.post(url_file, data=payload, files=files, timeout=30)
                        time.sleep(2)
        return success
    except requests.RequestException:
        return False

def calcular_nuevo_bloque() -> int:
    ms = int(time.time() * 1000)
    ultimo_digito = ms % 10
    return ultimo_digito if ultimo_digito > 0 else 5

def calcular_pausa_bloque() -> int:
    ms = int(time.time() * 1000)
    ultimos_dos = ms % 100
    return ultimos_dos + TIEMPO_MINIMO_BLOQUE if ultimos_dos < TIEMPO_MINIMO_BLOQUE else ultimos_dos

# ==============================================================================
# PROCESAMIENTO DE LOTE ESTRICTO
# ==============================================================================
def procesar_lote_envio(
    canal: str,
    instancia: str,
    token: str,
    contactos: List[Dict[str, Any]],
    plantilla: str,
    imagenes: List[str],
    timeout_wa_web: int = 18,
    callback_progreso: Optional[Any] = None,
    evento_cancelar: Optional[threading.Event] = None,
) -> List[Dict[str, str]]:
    total = len(contactos)
    tamano_bloque_actual = calcular_nuevo_bloque()
    contador_en_bloque = 0
    
    pendientes_fallidos: List[Dict[str, str]] = []
    primer_envio = True

    for i, c in enumerate(contactos):
        nombre = str(c.get("nombre", "")).strip() or "Amigo/a"
        numero = str(c.get("numero", "")).strip()

        if evento_cancelar and evento_cancelar.is_set():
            for restante in contactos[i:]:
                pendientes_fallidos.append({
                    "Nombre": str(restante.get("nombre", "")).strip() or "Amigo/a",
                    "Numero": str(restante.get("numero", "")).strip(),
                    "Motivo": "Campaña detenida antes de iniciar este contacto"
                })
            break

        if callback_progreso:
            callback_progreso(i, total, f"Enviando a {nombre} ({i + 1}/{total}) [Bloque: {contador_en_bloque + 1}/{tamano_bloque_actual}]...")

        texto = (
            plantilla.replace("{nombre}", nombre)
            .replace("{Nombre}", nombre)
            .replace("[nombre]", nombre)
            .replace("[Nombre]", nombre)
            .strip()
        )

        exito = False
        error_motivo = ""
        try:
            if canal == "WhatsApp Web (Gratis)":
                exito = dispatch_whatsapp_web_sin_recargas(
                    numero, texto, imagenes,
                    timeout_espera=timeout_wa_web,
                    es_primer_envio=primer_envio,
                    callback_progreso=callback_progreso,
                    evento_cancelar=evento_cancelar
                )
                if exito:
                    primer_envio = False
                else:
                    error_motivo = "No se pudo enviar el mensaje dentro del chat o el número es inválido"
            else:
                exito = dispatch_green_api(instancia, token, numero, texto, imagenes)
                if not exito:
                    error_motivo = "Green API rechazó la petición"
        except Exception as e:
            error_motivo = str(e)

        if not exito:
            pendientes_fallidos.append({
                "Nombre": nombre,
                "Numero": numero,
                "Motivo": error_motivo or "Fallo de envío"
            })
            if evento_cancelar and evento_cancelar.is_set():
                for restante in contactos[i + 1:]:
                    pendientes_fallidos.append({
                        "Nombre": str(restante.get("nombre", "")).strip() or "Amigo/a",
                        "Numero": str(restante.get("numero", "")).strip(),
                        "Motivo": "Campaña detenida por el usuario"
                    })
                break

        respaldar_mensaje(f"{nombre}_{numero}", texto, canal)
        registrar_historial({
            "canal": canal,
            "destinatario": f"{nombre} ({numero})",
            "mensaje": texto,
            "adjunto": f"{len(imagenes)} img" if imagenes else "No",
            "fecha_envio": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "estado": "Enviado" if exito else "Fallido",
        })

        if callback_progreso:
            callback_progreso(i + 1, total, f"Procesado {nombre} ({i + 1}/{total})")

        if i == total - 1:
            break

        contador_en_bloque += 1

        if contador_en_bloque >= tamano_bloque_actual:
            pausa_bloque = calcular_pausa_bloque()
            if callback_progreso:
                callback_progreso(i + 1, total, f"Pausa anti-spam entre bloques ({pausa_bloque}s)...")
            
            if not esperar_con_deteccion_mouse(pausa_bloque, callback_progreso, evento_cancelar):
                for restante in contactos[i + 1:]:
                    pendientes_fallidos.append({
                        "Nombre": str(restante.get("nombre", "")).strip() or "Amigo/a",
                        "Numero": str(restante.get("numero", "")).strip(),
                        "Motivo": "Cancelado durante pausa de seguridad"
                    })
                break

            tamano_bloque_actual = calcular_nuevo_bloque()
            contador_en_bloque = 0
        else:
            ms_indiv = int(time.time() * 1000) % 10
            pausa_individual = 10 + ms_indiv
            if not esperar_con_deteccion_mouse(pausa_individual, callback_progreso, evento_cancelar):
                for restante in contactos[i + 1:]:
                    pendientes_fallidos.append({
                        "Nombre": str(restante.get("nombre", "")).strip() or "Amigo/a",
                        "Numero": str(restante.get("numero", "")).strip(),
                        "Motivo": "Cancelado durante pausa individual"
                    })
                break

    if callback_progreso:
        if evento_cancelar and evento_cancelar.is_set():
            callback_progreso(0, total, "Campaña detenida.")
        else:
            callback_progreso(total, total, "¡Envíos completados!")

    return pendientes_fallidos

# ==============================================================================
# 6. CAPA DE MODELOS LLM
# ==============================================================================
def solicitar_completado_ia(proveedor: str, api_key: str, modelo: str, mensajes: List[Dict[str, str]]) -> str:
    if proveedor == "Demo / Offline":
        return "[Demo]: Texto procesado en modo local."
    if not api_key:
        return "Falta configurar la API Key en Ajustes."

    try:
        if proveedor == "Groq":
            if Groq is None:
                return "Instala: pip install groq"
            client = Groq(api_key=api_key)
            candidates = [modelo, "openai/gpt-oss-120b", "qwen/qwen3.8-27b"]
            for m in candidates:
                if not m:
                    continue
                try:
                    chat = client.chat.completions.create(messages=mensajes, model=m)
                    return chat.choices[0].message.content
                except Exception:
                    continue
            return "Error de conexión con Groq."

        if proveedor == "OpenAI":
            if OpenAI is None:
                return "Instala: pip install openai"
            client = OpenAI(api_key=api_key)
            chat = client.chat.completions.create(messages=mensajes, model=modelo or MODELOS_IA["OpenAI"])
            return chat.choices[0].message.content

        if proveedor == "Anthropic (Claude)":
            if anthropic is None:
                return "Instala: pip install anthropic"
            client = anthropic.Anthropic(api_key=api_key)
            sys_msg = next((m["content"] for m in mensajes if m["role"] == "system"), "")
            usr_msgs = [m for m in mensajes if m["role"] != "system"]
            msg = client.messages.create(
                model=modelo or MODELOS_IA["Anthropic (Claude)"],
                max_tokens=450,
                system=sys_msg,
                messages=usr_msgs,
            )
            return msg.content[0].text

        return f"Proveedor '{proveedor}' no reconocido."
    except Exception as e:
        return f"Error IA: {e}"

# ==============================================================================
# 7. INTERFAZ ULTRA-MINIMALISTA WAITSENDER
# ==============================================================================
class WaitSenderApp:
    def __init__(self, root: tk.Tk | TkinterDnD.Tk) -> None:
        self.root = root
        self.root.title("WaitSender - Powered by Teinek Solutions")
        self.root.geometry("1120x840")

        self.config: Dict[str, Any] = cargar_configuracion()
        self.repositorio_bases: Dict[str, List[Dict[str, Any]]] = cargar_repositorio_bases()
        self.base_activa_nombre: str = list(self.repositorio_bases.keys())[0] if self.repositorio_bases else "Base Principal"
        
        self.contactos_excel_temp: List[Dict[str, str]] = []
        self.imagenes_adjuntas: List[str] = []
        self.miniaturas_memoria: List[Any] = []

        self.chat_historial_ia: List[Dict[str, str]] = []

        self.cancel_event = threading.Event()
        self.is_recording = False
        self.stop_recording_event = threading.Event()

        self.scheduler: BackgroundScheduler = BackgroundScheduler()
        self.scheduler.start()

        self.current_theme_key: str = self.config.get("tema_actual", "dark")
        self.theme: Dict[str, str] = THEMES.get(self.current_theme_key, THEMES["dark"])

        self.estrella_image_ref: Optional[Any] = None
        self.registered_entry_widgets: List[tk.Entry] = []
        self.registered_text_widgets: List[tk.Text] = []

        self.ultimo_widget_enfocado: Optional[tk.Widget] = None

        self.construir_encabezado()
        self.construir_pestañas()
        self.aplicar_tema(self.current_theme_key)

    def cambiar_tema(self, nuevo_tema: str) -> None:
        if nuevo_tema not in THEMES:
            return
        self.current_theme_key = nuevo_tema
        self.theme = THEMES[nuevo_tema]
        self.config["tema_actual"] = nuevo_tema
        guardar_configuracion(self.config)
        self.aplicar_tema(nuevo_tema)

    def aplicar_tema(self, key: str) -> None:
        t = THEMES[key]
        self.root.configure(bg=t["bg_main"])
        self.header_frame.configure(bg=t["bg_main"])
        if hasattr(self, "lbl_estrella") and self.lbl_estrella:
            self.lbl_estrella.configure(bg=t["bg_main"])
        self.header_titles_frame.configure(bg=t["bg_main"])
        self.header_title_label.configure(bg=t["bg_main"], fg=t["text_main"])
        self.header_subtitle_label.configure(bg=t["bg_main"], fg=t["accent"])
        self.header_theme_selector_frame.configure(bg=t["bg_main"])

        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TNotebook", background=t["bg_main"], borderwidth=0)
        style.configure("TNotebook.Tab", background=t["bg_card"], foreground=t["text_muted"],
                        font=(t["font_family"], 9, "bold"), padding=[14, 6], borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", t["primary"])], foreground=[("selected", "#FFFFFF")])

        style.configure("TFrame", background=t["bg_main"], borderwidth=0)
        style.configure("Card.TFrame", background=t["bg_card"], relief="flat", borderwidth=0)
        style.configure("Inner.TFrame", background=t["bg_card"], borderwidth=0)
        style.configure("TLabel", background=t["bg_card"], foreground=t["text_main"], font=(t["font_family"], 9))
        style.configure("Header.TLabel", background=t["bg_card"], foreground=t["text_main"], font=(t["font_family"], 9, "bold"))
        style.configure("Sub.TLabel", background=t["bg_card"], foreground=t["text_muted"], font=(t["font_family"], 8))
        style.configure("Path.TLabel", background=t["bg_card"], foreground=t["accent"], font=(t["font_family"], 8, "italic"))

        style.configure("Primary.TButton", background=t["primary"], foreground="#FFFFFF",
                        font=(t["font_family"], 9, "bold"), borderwidth=0, padding=[10, 5])
        style.map("Primary.TButton", background=[("active", t["primary_hover"])])

        style.configure("Action.TButton", background=t["bg_input"], foreground=t["text_main"],
                        font=(t["font_family"], 8, "bold"), borderwidth=0, padding=[6, 4])

        style.configure("Treeview", background=t["bg_card"], foreground=t["text_main"],
                        fieldbackground=t["bg_card"], rowheight=26, font=(t["font_family"], 9), borderwidth=0)
        style.configure("Treeview.Heading", background=t["bg_input"], foreground=t["text_main"],
                        font=(t["font_family"], 9, "bold"), borderwidth=0)
        style.map("Treeview", background=[("selected", t["primary"])])

        for entry in self.registered_entry_widgets:
            entry.configure(bg=t["bg_input"], fg=t["text_main"], insertbackground=t["text_main"], relief="flat", highlightthickness=0)
        for text in self.registered_text_widgets:
            text.configure(bg=t["bg_input"], fg=t["text_main"], insertbackground=t["text_main"], relief="flat", highlightthickness=0)

        self.btn_ahora.configure(bg=t["btn_now"], activebackground=t["btn_now_hover"])
        self.btn_programar.configure(bg=t["primary"], activebackground=t["primary_hover"])
        self.actualizar_apariencia_boton_dictado()
        self.refrescar_galeria_recursos()

    def construir_encabezado(self) -> None:
        self.header_frame = tk.Frame(self.root, bg=self.theme["bg_main"], pady=6, padx=14)
        self.header_frame.pack(fill=tk.X)

        self.lbl_estrella: Optional[tk.Label] = None
        if PIL_AVAILABLE and os.path.exists(FILE_ESTRELLA):
            try:
                pil_img = Image.open(FILE_ESTRELLA)
                orig_w, orig_h = pil_img.size
                target_h = 36
                ratio = target_h / orig_h
                new_w = int(orig_w * ratio)
                resized = pil_img.resize((new_w, target_h), Image.Resampling.LANCZOS)
                self.estrella_image_ref = ImageTk.PhotoImage(resized)
                self.root.iconphoto(False, self.estrella_image_ref)

                self.lbl_estrella = tk.Label(self.header_frame, image=self.estrella_image_ref, bg=self.theme["bg_main"])
                self.lbl_estrella.pack(side=tk.LEFT, padx=(0, 8))
            except Exception:
                pass

        self.header_titles_frame = tk.Frame(self.header_frame, bg=self.theme["bg_main"])
        self.header_titles_frame.pack(side=tk.LEFT)

        self.header_title_label = tk.Label(
            self.header_titles_frame, text="WaitSender", font=(FONT_FAMILY, 14, "bold"),
            fg=self.theme["text_main"], bg=self.theme["bg_main"]
        )
        self.header_title_label.pack(anchor=tk.W)

        self.header_subtitle_label = tk.Label(
            self.header_titles_frame, text="Powered by Teinek Solutions",
            font=(FONT_FAMILY, 8), fg=self.theme["accent"], bg=self.theme["bg_main"]
        )
        self.header_subtitle_label.pack(anchor=tk.W)

        self.header_theme_selector_frame = tk.Frame(self.header_frame, bg=self.theme["bg_main"])
        self.header_theme_selector_frame.pack(side=tk.RIGHT, pady=2)

        for name, key in [("☀️ Claro", "light"), ("🌙 Oscuro", "dark"), ("👓 Filtro", "blue_filter")]:
            tk.Button(
                self.header_theme_selector_frame, text=name, font=(FONT_FAMILY, 8),
                bg=self.theme["bg_card"], fg=self.theme["text_main"], relief="flat", padx=6, pady=2,
                command=lambda k=key: self.cambiar_tema(k)
            ).pack(side=tk.LEFT, padx=2)

    def construir_pestañas(self) -> None:
        notebook = ttk.Notebook(self.root)
        self.tab_envios = ttk.Frame(notebook)
        self.tab_contactos = ttk.Frame(notebook)
        self.tab_recursos = ttk.Frame(notebook)
        self.tab_perfil = ttk.Frame(notebook)
        self.tab_historial = ttk.Frame(notebook)
        self.tab_config = ttk.Frame(notebook)

        notebook.add(self.tab_envios, text="  🚀 Envíos  ")
        notebook.add(self.tab_contactos, text="  👥 Bases & Categorías  ")
        notebook.add(self.tab_recursos, text="  🎨 Galería  ")
        notebook.add(self.tab_perfil, text="  👤 Base IA  ")
        notebook.add(self.tab_historial, text="  📊 Historial  ")
        notebook.add(self.tab_config, text="  ⚙️ Ajustes  ")
        notebook.pack(expand=True, fill="both", padx=12, pady=(2, 10))

        self.init_tab_contactos()
        self.init_tab_envios()
        self.init_tab_recursos()
        self.init_tab_perfil()
        self.init_tab_historial()
        self.init_tab_config()

    def init_tab_envios(self) -> None:
        f = ttk.Frame(self.tab_envios)
        f.pack(fill=tk.BOTH, expand=True, pady=4)

        card_switch = ttk.Frame(f, style="Card.TFrame", padding=6)
        card_switch.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(card_switch, text="Enviar a:", style="Header.TLabel").pack(side=tk.LEFT, padx=6)
        self.var_origen_envio = tk.StringVar(value="excel")

        for label, val in [
            ("📁 Excel", "excel"),
            ("👥 Base de Datos", "bd"),
            ("👤 Manual", "manual")
        ]:
            rb = tk.Radiobutton(
                card_switch, text=label, variable=self.var_origen_envio,
                value=val, bg=self.theme["bg_card"], fg=self.theme["text_main"],
                selectcolor=self.theme["bg_input"], activebackground=self.theme["bg_card"],
                relief="flat", highlightthickness=0,
                command=self.alternar_vistas_contextuales
            )
            rb.pack(side=tk.LEFT, padx=6)

        self.lbl_estado_origen = ttk.Label(card_switch, text="0 contactos listos", style="Path.TLabel")
        self.lbl_estado_origen.pack(side=tk.RIGHT, padx=6)

        self.card_contextual = ttk.Frame(f, style="Card.TFrame", padding=6)
        self.card_contextual.pack(fill=tk.X, pady=2)

        self.frame_ctx_excel = ttk.Frame(self.card_contextual, style="Inner.TFrame")
        ttk.Button(self.frame_ctx_excel, text="📁 Cargar Excel", style="Primary.TButton", command=self.importar_excel_reemplazo).pack(side=tk.LEFT, padx=4)
        self.lbl_excel_estado = ttk.Label(self.frame_ctx_excel, text="Sin archivo cargado", style="Sub.TLabel")
        self.lbl_excel_estado.pack(side=tk.LEFT, padx=6)

        self.frame_ctx_bd = ttk.Frame(self.card_contextual, style="Inner.TFrame")
        ttk.Label(self.frame_ctx_bd, text="Categoría:").pack(side=tk.LEFT, padx=4)
        self.combo_filtro_cat = ttk.Combobox(self.frame_ctx_bd, state="readonly", width=14)
        self.combo_filtro_cat.pack(side=tk.LEFT, padx=4)
        self.combo_filtro_cat.bind("<<ComboboxSelected>>", lambda e: self.actualizar_resumen_destinatarios())

        self.frame_ctx_manual = ttk.Frame(self.card_contextual, style="Inner.TFrame")
        ttk.Label(self.frame_ctx_manual, text="Número:").pack(side=tk.LEFT, padx=4)
        self.entry_manual_num = tk.Entry(self.frame_ctx_manual, relief="flat", width=16)
        self.entry_manual_num.pack(side=tk.LEFT, padx=4)
        self.entry_manual_num.bind("<KeyRelease>", lambda e: self.actualizar_resumen_destinatarios())
        self.registered_entry_widgets.append(self.entry_manual_num)

        ttk.Label(self.frame_ctx_manual, text="Nombre:").pack(side=tk.LEFT, padx=4)
        self.entry_manual_nom = tk.Entry(self.frame_ctx_manual, relief="flat", width=14)
        self.entry_manual_nom.pack(side=tk.LEFT, padx=4)
        self.registered_entry_widgets.append(self.entry_manual_nom)

        card_comun = ttk.Frame(f, style="Card.TFrame", padding=6)
        card_comun.pack(fill=tk.X, pady=(2, 2))

        ttk.Label(card_comun, text="Canal:").pack(side=tk.LEFT, padx=4)
        self.combo_canal = ttk.Combobox(card_comun, values=["WhatsApp Web (Gratis)", "Green API (Servidor)"], state="readonly", width=18)
        self.combo_canal.set(self.config.get("canal_predeterminado", "WhatsApp Web (Gratis)"))
        self.combo_canal.pack(side=tk.LEFT, padx=4)

        target_time = datetime.now() + timedelta(minutes=2)
        ttk.Label(card_comun, text="Programar:").pack(side=tk.LEFT, padx=(12, 4))
        self.entry_fecha = tk.Entry(card_comun, relief="flat", width=10)
        self.entry_fecha.insert(0, target_time.strftime("%Y-%m-%d"))
        self.entry_fecha.pack(side=tk.LEFT, padx=2)
        self.registered_entry_widgets.append(self.entry_fecha)

        self.entry_hora = tk.Entry(card_comun, relief="flat", width=6)
        self.entry_hora.insert(0, target_time.strftime("%H:%M"))
        self.entry_hora.pack(side=tk.LEFT, padx=2)
        self.registered_entry_widgets.append(self.entry_hora)

        ttk.Button(card_comun, text="📇 Exportar a .VCF", style="Action.TButton", command=self.generar_vcf_rapido).pack(side=tk.RIGHT, padx=4)

        card_central = ttk.Frame(f, style="Card.TFrame", padding=8)
        card_central.pack(fill=tk.BOTH, expand=True, pady=3)

        col_left = ttk.Frame(card_central, style="Inner.TFrame")
        col_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        ttk.Label(col_left, text="Mensaje Final a Enviar:", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 2))

        self.text_editor = tk.Text(col_left, height=8, relief="flat", wrap="word", font=(FONT_FAMILY, 9), padx=8, pady=8)
        empresa_actual = self.config.get("perfil_empresa", "Teinek Solutions")
        self.text_editor.insert(tk.END, f"Hola {{nombre}}, un cordial saludo por parte de {empresa_actual}.")
        self.text_editor.pack(fill=tk.BOTH, expand=True, pady=2)
        self.text_editor.bind("<FocusIn>", lambda e: self._registrar_foco(self.text_editor))
        self.registered_text_widgets.append(self.text_editor)

        f_tools = ttk.Frame(col_left, style="Inner.TFrame")
        f_tools.pack(fill=tk.X, pady=2)

        for label, token in [("B", "*"), ("I", "_"), ("S", "~")]:
            tk.Button(f_tools, text=label, font=(FONT_FAMILY, 8, "bold"), bg=self.theme["bg_input"],
                      fg=self.theme["text_main"], relief="flat", padx=5, pady=1,
                      command=lambda t=token: self.formatear_seleccion(t)).pack(side=tk.LEFT, padx=1)

        self.combo_emojis = ttk.Combobox(f_tools, values=["😀", "👋", "🚀", "✨", "💼", "📅", "✅", "⚠️", "📌"], state="readonly", width=3)
        self.combo_emojis.set("😀")
        self.combo_emojis.pack(side=tk.LEFT, padx=3)
        tk.Button(f_tools, text="➕", font=(FONT_FAMILY, 8), bg=self.theme["bg_input"], fg=self.theme["text_main"],
                  relief="flat", padx=4, pady=1, command=self.insertar_emoji).pack(side=tk.LEFT)

        tk.Button(f_tools, text="{nombre}", font=(FONT_FAMILY, 8, "italic"), bg=self.theme["bg_input"],
                  fg=self.theme["accent"], relief="flat", padx=5, pady=1, command=self.insertar_variable_nombre).pack(side=tk.LEFT, padx=3)

        self.btn_pulir_editor = tk.Button(
            f_tools, text="✨ Pulir", font=(FONT_FAMILY, 8, "bold"),
            bg=self.theme["primary"], fg="#FFFFFF", relief="flat", padx=8, pady=2,
            command=self.refinar_mensaje_final
        )
        self.btn_pulir_editor.pack(side=tk.RIGHT, padx=2)

        self.lbl_imgs_resumen = ttk.Label(col_left, text="0 fotos adjuntas (gestionar en Galería)", style="Sub.TLabel")
        self.lbl_imgs_resumen.pack(anchor=tk.W, pady=(2, 0))

        col_right = ttk.Frame(card_central, style="Inner.TFrame")
        col_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))

        ttk.Label(col_right, text="Asistente IA (Iterar mensaje):", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 2))

        self.text_sugerencia = tk.Text(col_right, height=6, relief="flat", wrap="word", font=(FONT_FAMILY, 9), padx=8, pady=8)
        self.text_sugerencia.pack(fill=tk.BOTH, expand=True, pady=2)
        self.registered_text_widgets.append(self.text_sugerencia)

        f_chat_ia = ttk.Frame(col_right, style="Inner.TFrame")
        f_chat_ia.pack(fill=tk.X, pady=2)

        self.entry_prompt = tk.Entry(f_chat_ia, relief="flat")
        self.entry_prompt.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.entry_prompt.bind("<FocusIn>", lambda e: self._registrar_foco(self.entry_prompt))
        self.entry_prompt.bind("<Return>", lambda e: self.enviar_instruccion_ia())
        self.registered_entry_widgets.append(self.entry_prompt)

        self.btn_enviar_ia = tk.Button(
            f_chat_ia, text="💬 Mejorar", font=(FONT_FAMILY, 8, "bold"),
            bg=self.theme["primary"], fg="#FFFFFF", relief="flat", padx=8, pady=2,
            command=self.enviar_instruccion_ia
        )
        self.btn_enviar_ia.pack(side=tk.RIGHT)

        f_ia_actions = ttk.Frame(col_right, style="Inner.TFrame")
        f_ia_actions.pack(fill=tk.X, pady=2)

        ttk.Button(f_ia_actions, text="🧹 Limpiar", style="Action.TButton", command=self.reiniciar_chat_ia).pack(side=tk.LEFT)
        ttk.Button(f_ia_actions, text="⬇️ Usar", style="Action.TButton", command=self.transferir_sugerencia).pack(side=tk.RIGHT)

        f_bottom = ttk.Frame(f, style="Card.TFrame", padding=6)
        f_bottom.pack(fill=tk.X, pady=3)

        self.btn_dictado_toggle = tk.Button(
            f_bottom, text="🎙️ Dictar", font=(FONT_FAMILY, 9, "bold"),
            bg=self.theme["accent"], fg="#000000", relief="flat", padx=12, pady=7,
            command=self.toggle_dictado
        )
        self.btn_dictado_toggle.pack(side=tk.LEFT, padx=4)

        self.btn_ahora = tk.Button(
            f_bottom, text="⚡ ENVIAR AHORA", font=(FONT_FAMILY, 10, "bold"),
            relief="flat", pady=8, fg="#FFFFFF", command=self.ejecutar_despacho_inmediato
        )
        self.btn_ahora.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        self.btn_programar = tk.Button(
            f_bottom, text="🚀 PROGRAMAR", font=(FONT_FAMILY, 10, "bold"),
            relief="flat", pady=8, fg="#FFFFFF", command=self.agendar_campana
        )
        self.btn_programar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=4)

        self.actualizar_selectores_categorias()
        self.alternar_vistas_contextuales()

    def alternar_vistas_contextuales(self) -> None:
        self.frame_ctx_excel.pack_forget()
        self.frame_ctx_bd.pack_forget()
        self.frame_ctx_manual.pack_forget()

        modo = self.var_origen_envio.get()
        if modo == "excel":
            self.frame_ctx_excel.pack(fill=tk.X)
        elif modo == "bd":
            self.frame_ctx_bd.pack(fill=tk.X)
        elif modo == "manual":
            self.frame_ctx_manual.pack(fill=tk.X)

        self.actualizar_resumen_destinatarios()

    def init_tab_contactos(self) -> None:
        f = ttk.Frame(self.tab_contactos, style="Card.TFrame", padding=12)
        f.pack(fill=tk.BOTH, expand=True, pady=4)

        f_grupo = ttk.Frame(f, style="Inner.TFrame")
        f_grupo.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(f_grupo, text="Base de Datos:", style="Header.TLabel").pack(side=tk.LEFT, padx=4)
        self.combo_bases_repo = ttk.Combobox(f_grupo, state="readonly", width=18)
        self.combo_bases_repo.pack(side=tk.LEFT, padx=4)
        self.combo_bases_repo.bind("<<ComboboxSelected>>", self.on_cambio_base_activa)

        ttk.Button(f_grupo, text="➕ Nueva BD", style="Action.TButton", command=self.crear_nueva_base).pack(side=tk.LEFT, padx=4)
        ttk.Button(f_grupo, text="🗑️ Eliminar BD", style="Action.TButton", command=self.eliminar_base_actual).pack(side=tk.LEFT, padx=4)
        ttk.Button(f_grupo, text="📇 Exportar a .VCF", style="Action.TButton", command=self.generar_vcf_de_bd).pack(side=tk.RIGHT, padx=4)

        f_input = ttk.Frame(f, style="Inner.TFrame")
        f_input.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(f_input, text="Nombre:").pack(side=tk.LEFT, padx=2)
        self.ent_nuevo_nombre = tk.Entry(f_input, relief="flat", width=14)
        self.ent_nuevo_nombre.pack(side=tk.LEFT, padx=2)
        self.registered_entry_widgets.append(self.ent_nuevo_nombre)

        ttk.Label(f_input, text="Teléfono:").pack(side=tk.LEFT, padx=2)
        self.ent_nuevo_numero = tk.Entry(f_input, relief="flat", width=15)
        self.ent_nuevo_numero.pack(side=tk.LEFT, padx=2)
        self.registered_entry_widgets.append(self.ent_nuevo_numero)

        ttk.Label(f_input, text="Categoría:").pack(side=tk.LEFT, padx=2)
        self.combo_categoria_input = ttk.Combobox(f_input, state="readonly", width=12)
        self.combo_categoria_input.pack(side=tk.LEFT, padx=2)

        ttk.Button(f_input, text="➕ Cat.", style="Action.TButton", command=self.crear_nueva_categoria).pack(side=tk.LEFT, padx=2)
        ttk.Button(f_input, text="💾 Agregar", style="Primary.TButton", command=self.agregar_contacto_bd).pack(side=tk.LEFT, padx=4)

        ttk.Button(f_input, text="📥 Cargar a BD", style="Action.TButton", command=self.importar_excel_a_bd).pack(side=tk.RIGHT, padx=4)
        ttk.Button(f_input, text="🗑️ Eliminar", style="Action.TButton", command=self.eliminar_contacto_seleccionado).pack(side=tk.RIGHT, padx=4)

        cols = ("sel", "nombre", "numero", "categoria")
        self.tree_contactos = ttk.Treeview(f, columns=cols, show="headings")
        self.tree_contactos.heading("sel", text="Activo")
        self.tree_contactos.heading("nombre", text="Nombre")
        self.tree_contactos.heading("numero", text="Teléfono")
        self.tree_contactos.heading("categoria", text="Categoría")
        
        self.tree_contactos.column("sel", width=65, anchor="center")
        self.tree_contactos.column("nombre", width=260)
        self.tree_contactos.column("numero", width=220)
        self.tree_contactos.column("categoria", width=180)
        self.tree_contactos.pack(fill=tk.BOTH, expand=True, pady=4)

        self.tree_contactos.bind("<Button-1>", self.on_click_contacto)
        self.actualizar_combos_bases()
        self.recargar_tabla_contactos()

    def actualizar_combos_bases(self) -> None:
        nombres = list(self.repositorio_bases.keys())
        self.combo_bases_repo["values"] = nombres
        if self.base_activa_nombre in nombres:
            self.combo_bases_repo.set(self.base_activa_nombre)
        elif nombres:
            self.base_activa_nombre = nombres[0]
            self.combo_bases_repo.set(self.base_activa_nombre)

    def on_cambio_base_activa(self, event: Optional[tk.Event] = None) -> None:
        self.base_activa_nombre = self.combo_bases_repo.get()
        self.recargar_tabla_contactos()
        self.actualizar_resumen_destinatarios()

    def crear_nueva_base(self) -> None:
        modal = tk.Toplevel(self.root)
        modal.title("Nueva Base")
        modal.geometry("300x120")
        modal.configure(bg=self.theme["bg_main"])
        modal.grab_set()

        ttk.Label(modal, text="Nombre del grupo:").pack(pady=(12, 4))
        ent_nom = tk.Entry(modal, relief="flat", width=25)
        ent_nom.pack(pady=4)

        def guardar():
            nombre = ent_nom.get().strip()
            if nombre and nombre not in self.repositorio_bases:
                self.repositorio_bases[nombre] = []
                guardar_repositorio_bases(self.repositorio_bases)
                self.base_activa_nombre = nombre
                self.actualizar_combos_bases()
                self.recargar_tabla_contactos()
                modal.destroy()

        tk.Button(modal, text="Crear", bg=self.theme["primary"], fg="#FFFFFF", relief="flat", padx=10, command=guardar).pack(pady=6)

    def eliminar_base_actual(self) -> None:
        if len(self.repositorio_bases) <= 1:
            messagebox.showwarning("Aviso", "Debes mantener al menos una base de datos.")
            return
        if messagebox.askyesno("Confirmar", f"¿Deseas eliminar la base '{self.base_activa_nombre}'?"):
            del self.repositorio_bases[self.base_activa_nombre]
            guardar_repositorio_bases(self.repositorio_bases)
            self.base_activa_nombre = list(self.repositorio_bases.keys())[0]
            self.actualizar_combos_bases()
            self.recargar_tabla_contactos()
            self.actualizar_resumen_destinatarios()

    def actualizar_selectores_categorias(self) -> None:
        cats = self.config.get("categorias_bd", ["General", "Clientes", "Prospectos", "Cobranza"])
        if hasattr(self, "combo_categoria_input"):
            self.combo_categoria_input["values"] = cats
            if cats:
                self.combo_categoria_input.set(cats[0])

        if hasattr(self, "combo_filtro_cat"):
            self.combo_filtro_cat["values"] = ["Todas"] + cats
            self.combo_filtro_cat.set("Todas")

    def crear_nueva_categoria(self) -> None:
        modal = tk.Toplevel(self.root)
        modal.title("Nueva Categoría")
        modal.geometry("280x110")
        modal.configure(bg=self.theme["bg_main"])
        modal.grab_set()

        ttk.Label(modal, text="Nombre:").pack(pady=(10, 4))
        ent_cat = tk.Entry(modal, relief="flat", width=22)
        ent_cat.pack(pady=4)

        def guardar():
            cat = ent_cat.get().strip().capitalize()
            cats = self.config.get("categorias_bd", [])
            if cat and cat not in cats:
                cats.append(cat)
                self.config["categorias_bd"] = cats
                guardar_configuracion(self.config)
                self.actualizar_selectores_categorias()
                self.combo_categoria_input.set(cat)
                modal.destroy()

        tk.Button(modal, text="Guardar", bg=self.theme["primary"], fg="#FFFFFF", relief="flat", padx=8, command=guardar).pack(pady=6)

    def recargar_tabla_contactos(self) -> None:
        for r in self.tree_contactos.get_children():
            self.tree_contactos.delete(r)
        
        contactos = self.repositorio_bases.get(self.base_activa_nombre, [])
        for i, c in enumerate(contactos):
            check_str = "✓ Sí" if c.get("activo", True) else "○ No"
            self.tree_contactos.insert("", tk.END, iid=str(i), values=(check_str, c.get("nombre", ""), c.get("numero", ""), c.get("categoria", "General")))

    def on_click_contacto(self, event: tk.Event) -> None:
        item_id = self.tree_contactos.identify_row(event.y)
        if item_id:
            idx = int(item_id)
            contactos = self.repositorio_bases.get(self.base_activa_nombre, [])
            contactos[idx]["activo"] = not contactos[idx].get("activo", True)
            guardar_repositorio_bases(self.repositorio_bases)
            self.recargar_tabla_contactos()
            self.actualizar_resumen_destinatarios()

    def agregar_contacto_bd(self) -> None:
        nom = self.ent_nuevo_nombre.get().strip() or "Amigo/a"
        num = validar_y_limpiar_telefono(self.ent_nuevo_numero.get().strip())
        cat = self.combo_categoria_input.get() or "General"
        if not num:
            messagebox.showwarning("Número Inválido", "Ingresa un teléfono válido de 10 a 15 dígitos.")
            return
        
        if self.base_activa_nombre not in self.repositorio_bases:
            self.repositorio_bases[self.base_activa_nombre] = []

        self.repositorio_bases[self.base_activa_nombre].append({
            "nombre": nom, "numero": num, "categoria": cat, "activo": True
        })
        guardar_repositorio_bases(self.repositorio_bases)
        self.recargar_tabla_contactos()
        self.actualizar_resumen_destinatarios()
        self.ent_nuevo_nombre.delete(0, tk.END)
        self.ent_nuevo_numero.delete(0, tk.END)

    def eliminar_contacto_seleccionado(self) -> None:
        sel = self.tree_contactos.selection()
        if not sel:
            return
        idx = int(sel[0])
        del self.repositorio_bases[self.base_activa_nombre][idx]
        guardar_repositorio_bases(self.repositorio_bases)
        self.recargar_tabla_contactos()
        self.actualizar_resumen_destinatarios()

    def importar_excel_a_bd(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if not path:
            return
        try:
            df = pd.read_excel(path)
            nuevos = self._extraer_contactos_de_df(df)
            cat_default = self.combo_categoria_input.get() or "General"
            for n in nuevos:
                n["categoria"] = cat_default
                n["activo"] = True

            self.repositorio_bases.setdefault(self.base_activa_nombre, []).extend(nuevos)
            guardar_repositorio_bases(self.repositorio_bases)
            self.recargar_tabla_contactos()
            self.actualizar_resumen_destinatarios()
            messagebox.showinfo("Excel", f"Se añadieron {len(nuevos)} contactos a '{self.base_activa_nombre}'.")
        except Exception as e:
            messagebox.showerror("Error", f"Fallo al leer Excel: {e}")

    def _extraer_contactos_de_df(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        lista = []
        cols = [str(c).lower().strip() for c in df.columns]
        idx_nom = 0
        for pos, col in enumerate(cols):
            if any(k in col for k in ["nom", "name", "destinatario", "cliente"]):
                idx_nom = pos
                break

        idx_tel = 1 if len(cols) > 1 else 0
        for pos, col in enumerate(cols):
            if any(k in col for k in ["tel", "phone", "cel", "movil", "numero", "wa"]):
                idx_tel = pos
                break

        for _, row in df.iterrows():
            nom = str(row.iloc[idx_nom]).strip()
            tel_raw = str(row.iloc[idx_tel]).replace(".0", "").strip()
            tel_valido = validar_y_limpiar_telefono(tel_raw)
            if tel_valido:
                lista.append({"nombre": nom if nom.lower() != "nan" and nom else "Amigo/a", "numero": tel_valido})
        return lista

    def importar_excel_reemplazo(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if not path:
            return
        try:
            df = pd.read_excel(path)
            self.contactos_excel_temp = self._extraer_contactos_de_df(df)
            nombre_archivo = os.path.basename(path)
            self.lbl_excel_estado.config(text=f"{nombre_archivo} ({len(self.contactos_excel_temp)} contactos)")
            self.var_origen_envio.set("excel")
            self.alternar_vistas_contextuales()
            messagebox.showinfo("Excel Cargado", f"Se cargaron {len(self.contactos_excel_temp)} contactos.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el Excel: {e}")

    def obtener_contactos_para_despacho(self) -> List[Dict[str, Any]]:
        origen = self.var_origen_envio.get()
        if origen == "excel":
            return list(self.contactos_excel_temp)
        elif origen == "bd":
            contactos = self.repositorio_bases.get(self.base_activa_nombre, [])
            filtro = self.combo_filtro_cat.get()
            if filtro == "Todas" or not filtro:
                return [c for c in contactos if c.get("activo", True)]
            return [c for c in contactos if c.get("activo", True) and c.get("categoria") == filtro]
        else:
            num = validar_y_limpiar_telefono(self.entry_manual_num.get().strip())
            nom = self.entry_manual_nom.get().strip() or "Amigo/a"
            return [{"nombre": nom, "numero": num}] if num else []

    def actualizar_resumen_destinatarios(self) -> None:
        lista = self.obtener_contactos_para_despacho()
        origen = self.var_origen_envio.get().upper()
        self.lbl_estado_origen.config(text=f"{len(lista)} contactos listos [{origen}]")

    def generar_vcf_rapido(self) -> None:
        contactos = self.obtener_contactos_para_despacho()
        if not contactos:
            messagebox.showwarning("Aviso", "No hay contactos activos para exportar a la libreta.")
            return
        ruta = exportar_contactos_a_vcf(contactos, f"contactos_{self.var_origen_envio.get()}.vcf")
        msg = (
            f"Se generó el archivo de contactos para tu cuenta de WhatsApp/Google:\n{ruta}\n\n"
            f"¿Deseas abrir la carpeta para importarlo a tu agenda?"
        )
        if messagebox.askyesno("Contactos Generados", msg):
            self.abrir_explorador(DIR_CONTACTOS_VCF)

    def generar_vcf_de_bd(self) -> None:
        contactos = self.repositorio_bases.get(self.base_activa_nombre, [])
        if not contactos:
            messagebox.showwarning("Aviso", "La base de datos actual está vacía.")
            return
        ruta = exportar_contactos_a_vcf(contactos, f"base_{self.base_activa_nombre}.vcf")
        if messagebox.askyesno("vCard Creado", f"Archivo generado en:\n{ruta}\n\n¿Abrir carpeta?"):
            self.abrir_explorador(DIR_CONTACTOS_VCF)

    def compilar_prompt_sistema(self) -> str:
        return (
            f"Eres el asistente redactor de {self.config.get('perfil_nombre', 'Nombre')} "
            f"en {self.config.get('perfil_empresa', 'Teinek Solutions')}.\n"
            f"DOCUMENTACIÓN:\n{self.txt_biblio.get('1.0', tk.END).strip()}\n\n"
            f"1. Si necesitas incluir el nombre de la persona usa {{nombre}}.\n"
            f"2. Formato de WhatsApp (*negrita* y emojis moderados).\n"
            f"3. Responde directamente con el mensaje sin preámbulos."
        )

    def refinar_mensaje_final(self) -> None:
        content = self.text_editor.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Aviso", "Escribe primero un mensaje en el cuadro principal.")
            return

        self.btn_pulir_editor.config(text="Puliendo...", state="disabled")

        def async_worker():
            prompt_pulido = (
                "Actúa como corrector de estilo para WhatsApp. Tareas estrictas:\n"
                "- Conectar ideas con mayor coherencia y fluidez.\n"
                "- Eliminar redundancias y palabras innecesarias.\n"
                "- Corregir ortografía, acentuación y puntuación.\n"
                "- Mantener el sentido, intenciones y tono original intactos.\n"
                "- Mantener la variable '{nombre}' sin alteraciones.\n"
                "NO inventes contenido ni cambies drásticamente el texto.\n"
                "Retorna ÚNICAMENTE el texto pulido:\n\n"
                f"{content}"
            )
            mensajes = [
                {"role": "system", "content": "Eres un editor de estilo y gramática quirúrgico."},
                {"role": "user", "content": prompt_pulido}
            ]
            res = solicitar_completado_ia(
                self.config.get("llm_provider"), self.config.get("llm_key"), self.config.get("llm_model"), mensajes
            )

            def aplicar():
                self.btn_pulir_editor.config(text="✨ Pulir", state="normal")
                if res and not res.startswith("Error"):
                    self.text_editor.delete("1.0", tk.END)
                    self.text_editor.insert(tk.END, res.strip())
                else:
                    messagebox.showerror("Error", res)

            self.root.after(0, aplicar)

        threading.Thread(target=async_worker, daemon=True).start()

    def enviar_instruccion_ia(self) -> None:
        instruccion = self.entry_prompt.get().strip()
        if not instruccion:
            return

        self.entry_prompt.delete(0, tk.END)
        self.btn_enviar_ia.config(text="...", state="disabled")

        def async_worker():
            if not self.chat_historial_ia:
                self.chat_historial_ia.append({"role": "system", "content": self.compilar_prompt_sistema()})

            self.chat_historial_ia.append({"role": "user", "content": instruccion})
            res = solicitar_completado_ia(
                self.config.get("llm_provider"), self.config.get("llm_key"), self.config.get("llm_model"), self.chat_historial_ia
            )

            def actualizar():
                self.btn_enviar_ia.config(text="💬 Mejorar", state="normal")
                if res and not res.startswith("Error"):
                    self.chat_historial_ia.append({"role": "assistant", "content": res.strip()})
                    self.text_sugerencia.delete("1.0", tk.END)
                    self.text_sugerencia.insert(tk.END, res.strip())
                else:
                    messagebox.showerror("Error IA", res)

            self.root.after(0, actualizar)

        threading.Thread(target=async_worker, daemon=True).start()

    def reiniciar_chat_ia(self) -> None:
        self.chat_historial_ia.clear()
        self.text_sugerencia.delete("1.0", tk.END)
        self.entry_prompt.delete(0, tk.END)

    def transferir_sugerencia(self) -> None:
        data = self.text_sugerencia.get("1.0", tk.END).strip()
        if data:
            self.text_editor.delete("1.0", tk.END)
            self.text_editor.insert(tk.END, data)

    def abrir_modal_envio(self, total: int) -> tuple[tk.Toplevel, ttk.Progressbar, tk.Label]:
        modal = tk.Toplevel(self.root)
        modal.title("Despacho - WaitSender")
        modal.geometry("440x180")
        modal.configure(bg=self.theme["bg_main"])
        modal.resizable(False, False)
        modal.grab_set()

        lbl_info = tk.Label(modal, text=f"Iniciando (0 de {total})...", font=(FONT_FAMILY, 9),
                            fg=self.theme["text_main"], bg=self.theme["bg_main"])
        lbl_info.pack(pady=(20, 8))

        prog_bar = ttk.Progressbar(modal, orient="horizontal", length=340, mode="determinate")
        prog_bar.pack(pady=8)

        btn_detener = tk.Button(
            modal, text="⏹️ Cancelar", font=(FONT_FAMILY, 9, "bold"),
            bg=self.theme["danger"], fg="#FFFFFF", relief="flat", padx=12, pady=5,
            command=self.cancelar_campana_actual
        )
        btn_detener.pack(pady=(10, 8))

        return modal, prog_bar, lbl_info

    def cancelar_campana_actual(self) -> None:
        self.cancel_event.set()

    def ejecutar_despacho_inmediato(self) -> None:
        contactos = self.obtener_contactos_para_despacho()
        if not contactos:
            messagebox.showwarning("Sin Destinatarios", "No hay contactos listos para enviar.")
            return

        msg = self.text_editor.get("1.0", tk.END).strip()
        if not msg:
            messagebox.showerror("Aviso", "El mensaje no puede estar vacío.")
            return

        canal = self.combo_canal.get()
        gid = self.config.get("green_id")
        gtok = self.config.get("green_token")
        if canal == "Green API (Servidor)" and (not gid or not gtok):
            messagebox.showerror("Error", "Configura las credenciales de Green API en Ajustes.")
            return

        self.cancel_event.clear()
        total = len(contactos)
        modal, prog_bar, lbl_info = self.abrir_modal_envio(total)
        timeout_wa = self.config.get("timeout_wa_web", 18)

        def actualizar_ui(actual: int, tot: int, estado: str):
            def f():
                if actual >= 0 and tot > 0:
                    porcentaje = (actual / tot * 100)
                    prog_bar["value"] = porcentaje
                    lbl_info.config(text=f"{estado} ({int(porcentaje)}%)")
                else:
                    lbl_info.config(text=estado)
            self.root.after(0, f)

        def cerrar_modal_y_reportar(fallidos: List[Dict[str, str]]):
            try:
                modal.grab_release()
                modal.destroy()
            except Exception:
                pass

            if fallidos:
                ruta_txt, ruta_xlsx = generar_reportes_dobles_fallos(fallidos)
                msg_reporte = (
                    f"La campaña concluyó con {len(fallidos)} contactos pendientes/no enviados.\n\n"
                    f"Archivos listos en 'reporte_fallos':\n"
                    f"• {os.path.basename(ruta_txt)} (Explicación del error)\n"
                    f"• {os.path.basename(ruta_xlsx)} (Excel listo para recargar y reintentar)\n\n"
                    f"¿Deseas abrir la carpeta de fallos ahora?"
                )
                if messagebox.askyesno("Reporte Completo Generado", msg_reporte):
                    self.abrir_explorador(DIR_FALLOS)
            else:
                if self.cancel_event.is_set():
                    messagebox.showwarning("Detenido", "La campaña fue cancelada.")
                else:
                    messagebox.showinfo("Completado", "¡Todos los mensajes fueron enviados!")

        def worker():
            fallidos: List[Dict[str, str]] = []
            try:
                fallidos = procesar_lote_envio(
                    canal, gid, gtok, contactos, msg, list(self.imagenes_adjuntas),
                    timeout_wa_web=timeout_wa,
                    callback_progreso=actualizar_ui,
                    evento_cancelar=self.cancel_event
                )
            except Exception as exc:
                fallidos = [{
                    "Nombre": str(c.get("nombre", "")).strip() or "Amigo/a",
                    "Numero": str(c.get("numero", "")).strip(),
                    "Motivo": f"Error imprevisto: {exc}"
                } for c in contactos]
            finally:
                self.root.after(0, lambda: cerrar_modal_y_reportar(fallidos))

        threading.Thread(target=worker, daemon=True).start()

    def agendar_campana(self) -> None:
        contactos = self.obtener_contactos_para_despacho()
        if not contactos:
            messagebox.showwarning("Aviso", "No hay contactos para programar.")
            return

        msg = self.text_editor.get("1.0", tk.END).strip()
        if not msg:
            messagebox.showerror("Aviso", "El mensaje no puede estar vacío.")
            return

        date_str = self.entry_fecha.get().strip()
        time_str = self.entry_hora.get().strip()
        try:
            target_dt = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
        except ValueError:
            messagebox.showerror("Formato Inválido", "Requiere fecha AAAA-MM-DD y hora HH:MM.")
            return

        timeout_wa = self.config.get("timeout_wa_web", 18)
        job_id = f"job_{int(target_dt.timestamp())}"
        self.scheduler.add_job(
            func=procesar_lote_envio,
            trigger="date",
            run_date=target_dt,
            args=[
                self.combo_canal.get(), self.config.get("green_id"), self.config.get("green_token"),
                contactos, msg, list(self.imagenes_adjuntas), timeout_wa
            ],
            id=job_id,
        )
        messagebox.showinfo("Agendado", f"Envío programado para {date_str} a las {time_str}.")

    def init_tab_recursos(self) -> None:
        f = ttk.Frame(self.tab_recursos, style="Card.TFrame", padding=10)
        f.pack(fill=tk.BOTH, expand=True, pady=2)

        f_top = ttk.Frame(f, style="Inner.TFrame")
        f_top.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(f_top, text="Haz clic para seleccionar imágenes:", style="Header.TLabel").pack(side=tk.LEFT)
        ttk.Button(f_top, text="📥 Importar", style="Primary.TButton", command=self.importar_imagenes_a_recursos).pack(side=tk.RIGHT, padx=3)
        ttk.Button(f_top, text="📂 Carpeta", style="Action.TButton", command=lambda: self.abrir_explorador(DIR_RECURSOS)).pack(side=tk.RIGHT, padx=3)

        self.canvas_galeria = tk.Canvas(f, bg=self.theme["bg_card"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(f, orient="vertical", command=self.canvas_galeria.yview)
        self.scroll_galeria = tk.Frame(self.canvas_galeria, bg=self.theme["bg_card"])

        self.scroll_galeria.bind("<Configure>", lambda e: self.canvas_galeria.configure(scrollregion=self.canvas_galeria.bbox("all")))
        self.canvas_galeria.create_window((0, 0), window=self.scroll_galeria, anchor="nw")
        self.canvas_galeria.configure(yscrollcommand=scrollbar.set)

        self.canvas_galeria.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        if DND_AVAILABLE:
            try:
                self.canvas_galeria.drop_target_register(DND_FILES)
                self.canvas_galeria.dnd_bind("<<Drop>>", self.on_drop_imagenes_galeria)
                self.scroll_galeria.drop_target_register(DND_FILES)
                self.scroll_galeria.drop_target_register(DND_FILES)
            except Exception:
                pass

    def on_drop_imagenes_galeria(self, event: tk.Event) -> None:
        valid_ext = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
        archivos = self._parsear_dnd(event.data)
        for f in archivos:
            if os.path.splitext(f)[1].lower() in valid_ext:
                dest = os.path.join(DIR_RECURSOS, os.path.basename(f))
                if os.path.abspath(f) != os.path.abspath(dest):
                    shutil.copy2(f, dest)
                if dest not in self.imagenes_adjuntas:
                    self.imagenes_adjuntas.append(dest)
        self.refrescar_galeria_recursos()

    def refrescar_galeria_recursos(self) -> None:
        for w in self.scroll_galeria.winfo_children():
            w.destroy()

        self.miniaturas_memoria.clear()
        valid_ext = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
        archivos = [os.path.join(DIR_RECURSOS, f) for f in os.listdir(DIR_RECURSOS) if os.path.splitext(f)[1].lower() in valid_ext]

        if not archivos:
            tk.Label(self.scroll_galeria, text="No hay imágenes en Galería.",
                     fg=self.theme["text_muted"], bg=self.theme["bg_card"]).pack(pady=30)
            self.lbl_imgs_resumen.config(text="0 fotos adjuntas")
            return

        row, col = 0, 0
        for path in archivos:
            seleccionada = path in self.imagenes_adjuntas
            borde_color = self.theme["accent"] if seleccionada else self.theme["bg_input"]

            card = tk.Frame(self.scroll_galeria, bg=borde_color, padx=3, pady=3, cursor="hand2")
            card.grid(row=row, column=col, padx=6, pady=6, sticky="nsew")

            inner_card = tk.Frame(card, bg=self.theme["bg_input"], padx=4, pady=4)
            inner_card.pack(fill=tk.BOTH, expand=True)

            if PIL_AVAILABLE:
                try:
                    img = Image.open(path)
                    img.thumbnail((110, 110), Image.Resampling.LANCZOS)
                    tk_img = ImageTk.PhotoImage(img)
                    self.miniaturas_memoria.append(tk_img)
                    lbl_img = tk.Label(inner_card, image=tk_img, bg=self.theme["bg_input"], cursor="hand2")
                    lbl_img.pack()
                    lbl_img.bind("<Button-1>", lambda e, p=path: self.toggle_seleccion_imagen(p))
                except Exception:
                    pass

            nom = os.path.basename(path)
            lbl_nom = tk.Label(inner_card, text=nom[:12] + (".." if len(nom) > 12 else ""), font=(FONT_FAMILY, 8),
                               fg=self.theme["text_main"], bg=self.theme["bg_input"], cursor="hand2")
            lbl_nom.pack(pady=(2, 0))
            lbl_nom.bind("<Button-1>", lambda e, p=path: self.toggle_seleccion_imagen(p))

            estado_txt = "✓ Sí" if seleccionada else "○ Añadir"
            lbl_estado = tk.Label(inner_card, text=estado_txt, font=(FONT_FAMILY, 7, "bold"),
                                  fg=self.theme["accent"] if seleccionada else self.theme["text_muted"],
                                  bg=self.theme["bg_input"], cursor="hand2")
            lbl_estado.pack()
            lbl_estado.bind("<Button-1>", lambda e, p=path: self.toggle_seleccion_imagen(p))

            card.bind("<Button-1>", lambda e, p=path: self.toggle_seleccion_imagen(p))
            inner_card.bind("<Button-1>", lambda e, p=path: self.toggle_seleccion_imagen(p))

            col += 1
            if col > 3:
                col = 0
                row += 1

        self.lbl_imgs_resumen.config(text=f"{len(self.imagenes_adjuntas)} foto(s) adjuntas")

    def toggle_seleccion_imagen(self, ruta: str) -> None:
        if ruta in self.imagenes_adjuntas:
            self.imagenes_adjuntas.remove(ruta)
        else:
            self.imagenes_adjuntas.append(ruta)
        self.refrescar_galeria_recursos()

    def importar_imagenes_a_recursos(self) -> None:
        files = filedialog.askopenfilenames(title="Importar Imágenes", filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.webp *.bmp")])
        if files:
            for f in files:
                dest = os.path.join(DIR_RECURSOS, os.path.basename(f))
                if os.path.abspath(f) != os.path.abspath(dest):
                    shutil.copy2(f, dest)
                if dest not in self.imagenes_adjuntas:
                    self.imagenes_adjuntas.append(dest)
            self.refrescar_galeria_recursos()

    def _registrar_foco(self, widget: tk.Widget) -> None:
        self.ultimo_widget_enfocado = widget

    def toggle_dictado(self) -> None:
        if self.is_recording:
            self.stop_recording_event.set()
        else:
            self.iniciar_dictado_universal()

    def actualizar_apariencia_boton_dictado(self) -> None:
        if self.is_recording:
            self.btn_dictado_toggle.config(text="⏹️ Grabando...", bg=self.theme["danger"], fg="#FFFFFF")
        else:
            self.btn_dictado_toggle.config(text="🎙️ Dictar", bg=self.theme["accent"], fg="#000000")

    def iniciar_dictado_universal(self) -> None:
        try:
            import sounddevice as sd
            from scipy.io.wavfile import write
        except ImportError:
            messagebox.showerror("Librerías Faltantes", "Instala en terminal:\npy -m pip install sounddevice numpy scipy")
            return

        self.is_recording = True
        self.stop_recording_event.clear()
        self.actualizar_apariencia_boton_dictado()
        destino = self.ultimo_widget_enfocado or self.text_editor

        def async_worker():
            sample_rate = 44100
            grabacion = []
            try:
                def callback(indata, frames, time_info, status):
                    grabacion.append(indata.copy())

                stream = sd.InputStream(samplerate=sample_rate, channels=1, dtype='int16', callback=callback)
                with stream:
                    seg = 0
                    while seg < 15 and not self.stop_recording_event.is_set():
                        time.sleep(0.4)
                        seg += 0.4

                if grabacion:
                    import numpy as np
                    audio_np = np.concatenate(grabacion, axis=0)
                    wav_io = io.BytesIO()
                    write(wav_io, sample_rate, audio_np)
                    wav_io.seek(0)

                    r = sr.Recognizer()
                    with sr.AudioFile(wav_io) as source:
                        audio_data = r.record(source)
                        texto = r.recognize_google(audio_data, language="es-MX")

                    def escribir():
                        if isinstance(destino, tk.Text):
                            destino.insert(tk.END, f" {texto}")
                        elif isinstance(destino, tk.Entry):
                            previo = destino.get()
                            destino.delete(0, tk.END)
                            destino.insert(0, f"{previo} {texto}".strip())

                    self.root.after(0, escribir)
            except Exception as exc:
                print(f"Aviso dictado: {exc}")
            finally:
                def reset():
                    self.is_recording = False
                    self.actualizar_apariencia_boton_dictado()
                self.root.after(0, reset)

        threading.Thread(target=async_worker, daemon=True).start()

    def init_tab_perfil(self) -> None:
        f = ttk.Frame(self.tab_perfil, style="Card.TFrame", padding=14)
        f.pack(fill=tk.BOTH, expand=True, pady=4)

        ttk.Label(f, text="Perfil y Base Documental IA:", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 4))
        f_dir = ttk.Frame(f, style="Inner.TFrame")
        f_dir.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(f_dir, text=f"📁 {DIR_CONOCIMIENTO}", style="Path.TLabel").pack(side=tk.LEFT)
        ttk.Button(f_dir, text="📂 Abrir", style="Action.TButton", command=lambda: self.abrir_explorador(DIR_CONOCIMIENTO)).pack(side=tk.RIGHT)

        grid = ttk.Frame(f, style="Inner.TFrame")
        grid.pack(fill=tk.X, pady=4)

        ttk.Label(grid, text="Nombre:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.ent_nombre = tk.Entry(grid, relief="flat", width=25)
        self.ent_nombre.insert(0, self.config.get("perfil_nombre", "Nombre"))
        self.ent_nombre.grid(row=0, column=1, padx=6, sticky=tk.W, pady=4)
        self.registered_entry_widgets.append(self.ent_nombre)

        ttk.Label(grid, text="Rol:").grid(row=0, column=2, sticky=tk.W, padx=(12, 0), pady=4)
        self.ent_rol = tk.Entry(grid, relief="flat", width=25)
        self.ent_rol.insert(0, self.config.get("perfil_rol", "Especialista en Automatización"))
        self.ent_rol.grid(row=0, column=3, padx=6, sticky=tk.W, pady=4)
        self.registered_entry_widgets.append(self.ent_rol)

        ttk.Label(grid, text="Empresa:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.ent_empresa = tk.Entry(grid, relief="flat", width=25)
        self.ent_empresa.insert(0, self.config.get("perfil_empresa", "Teinek Solutions"))
        self.ent_empresa.grid(row=1, column=1, padx=6, sticky=tk.W, pady=4)
        self.registered_entry_widgets.append(self.ent_empresa)

        f_doc_ctrl = ttk.Frame(f, style="Inner.TFrame")
        f_doc_ctrl.pack(fill=tk.X, pady=(10, 2))
        ttk.Label(f_doc_ctrl, text="Perfil / Bibliografía y Documentos de Contexto:", style="Header.TLabel").pack(side=tk.LEFT)
        ttk.Button(f_doc_ctrl, text="📥 Cargar PDF/Word", style="Action.TButton", command=self.examinar_documento).pack(side=tk.RIGHT)

        self.txt_biblio = tk.Text(f, height=8, relief="flat", wrap="word", padx=8, pady=8)
        self.txt_biblio.insert(tk.END, self.config.get("perfil_bibliografia", "Estilo profesional, claro y enfocado en soluciones ágiles."))
        self.txt_biblio.pack(fill=tk.BOTH, expand=True, pady=4)
        self.registered_text_widgets.append(self.txt_biblio)

        ttk.Button(f, text="💾 Guardar", style="Primary.TButton", command=self.salvar_perfil).pack(anchor=tk.E, pady=4)

    def init_tab_historial(self) -> None:
        f = ttk.Frame(self.tab_historial, style="Card.TFrame", padding=12)
        f.pack(fill=tk.BOTH, expand=True, pady=4)

        f_bar = ttk.Frame(f, style="Inner.TFrame")
        f_bar.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(f_bar, text="Registro General", style="Header.TLabel").pack(side=tk.LEFT)
        ttk.Button(f_bar, text="📂 Fallos (.txt y .xlsx)", style="Primary.TButton", command=lambda: self.abrir_explorador(DIR_FALLOS)).pack(side=tk.RIGHT, padx=4)
        ttk.Button(f_bar, text="📂 Historial", style="Action.TButton", command=lambda: self.abrir_explorador(DIR_HISTORIAL)).pack(side=tk.RIGHT, padx=4)

        cols = ("canal", "dest", "mensaje", "adjunto", "fecha", "estado")
        self.tree = ttk.Treeview(f, columns=cols, show="headings")
        self.tree.heading("canal", text="Canal")
        self.tree.heading("dest", text="Destinatario")
        self.tree.heading("mensaje", text="Mensaje")
        self.tree.heading("adjunto", text="Fotos")
        self.tree.heading("fecha", text="Fecha")
        self.tree.heading("estado", text="Estado")

        self.tree.column("canal", width=110)
        self.tree.column("dest", width=150)
        self.tree.column("mensaje", width=380)
        self.tree.column("adjunto", width=60)
        self.tree.column("fecha", width=120)
        self.tree.column("estado", width=70)
        self.tree.pack(fill=tk.BOTH, expand=True, pady=4)

        f_foot = ttk.Frame(f, style="Inner.TFrame")
        f_foot.pack(fill=tk.X)
        ttk.Button(f_foot, text="🔄 Refrescar", style="Action.TButton", command=self.recargar_tabla_historial).pack(side=tk.RIGHT)
        self.recargar_tabla_historial()

    def init_tab_config(self) -> None:
        f = ttk.Frame(self.tab_config, style="Card.TFrame", padding=14)
        f.pack(fill=tk.BOTH, expand=True, pady=4)

        ttk.Label(f, text="Ajustes de Plataforma", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 8))

        grid = ttk.Frame(f, style="Inner.TFrame")
        grid.pack(fill=tk.X, pady=4)

        ttk.Label(grid, text="Espera WA Web:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.spin_wa_timeout = tk.Spinbox(grid, from_=10, to=45, width=6, relief="flat")
        self.spin_wa_timeout.delete(0, "end")
        self.spin_wa_timeout.insert(0, str(self.config.get("timeout_wa_web", 18)))
        self.spin_wa_timeout.grid(row=0, column=1, padx=6, sticky=tk.W)
        ttk.Label(grid, text="segundos (18s seguro)", style="Sub.TLabel").grid(row=0, column=2, sticky=tk.W)

        ttk.Label(grid, text="Proveedor IA:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.combo_prov = ttk.Combobox(grid, values=["Groq", "OpenAI", "Anthropic (Claude)", "Demo / Offline"], state="readonly", width=20)
        self.combo_prov.set(self.config.get("llm_provider", "Groq"))
        self.combo_prov.grid(row=1, column=1, padx=6, sticky=tk.W, pady=4)

        ttk.Label(grid, text="API Key:").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.ent_key = tk.Entry(grid, relief="flat", width=40, show="*")
        self.ent_key.insert(0, self.config.get("llm_key", ""))
        self.ent_key.grid(row=2, column=1, padx=6, sticky=tk.W, pady=4)
        self.registered_entry_widgets.append(self.ent_key)

        ttk.Label(grid, text="Green API ID:").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.ent_gid = tk.Entry(grid, relief="flat", width=22)
        self.ent_gid.insert(0, self.config.get("green_id", ""))
        self.ent_gid.grid(row=3, column=1, padx=6, sticky=tk.W, pady=4)
        self.registered_entry_widgets.append(self.ent_gid)

        ttk.Label(grid, text="Green Token:").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.ent_gtok = tk.Entry(grid, relief="flat", width=40, show="*")
        self.ent_gtok.insert(0, self.config.get("green_token", ""))
        self.ent_gtok.grid(row=4, column=1, padx=6, sticky=tk.W, pady=4)
        self.registered_entry_widgets.append(self.ent_gtok)

        ttk.Button(f, text="💾 Guardar", style="Primary.TButton", command=self.salvar_credenciales).pack(anchor=tk.E, pady=12)

    def formatear_seleccion(self, token: str) -> None:
        try:
            sel_text = self.text_editor.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.text_editor.delete(tk.SEL_FIRST, tk.SEL_LAST)
            self.text_editor.insert(tk.INSERT, f"{token}{sel_text}{token}")
        except tk.TclError:
            self.text_editor.insert(tk.INSERT, f"{token}texto{token}")

    def insertar_emoji(self) -> None:
        char = self.combo_emojis.get()
        if char:
            self.text_editor.insert(tk.INSERT, f" {char} ")

    def insertar_variable_nombre(self) -> None:
        self.text_editor.insert(tk.INSERT, "{nombre}")

    def abrir_explorador(self, folder: str) -> None:
        if os.path.exists(folder):
            os.startfile(folder)

    def _parsear_dnd(self, data_str: str) -> List[str]:
        items = []
        if "{" in data_str:
            for piece in data_str.split("}"):
                c = piece.replace("{", "").strip()
                if c and os.path.exists(c):
                    items.append(c)
        else:
            for piece in data_str.split():
                if os.path.exists(piece):
                    items.append(piece)
        return items

    def examinar_documento(self) -> None:
        file = filedialog.askopenfilename(filetypes=[("Documentos", "*.pdf *.docx *.xlsx *.txt")])
        if file:
            self.procesar_archivo_contexto(file)

    def procesar_archivo_contexto(self, source_path: str) -> None:
        text_data = ""
        ext = os.path.splitext(source_path)[1].lower()
        name = os.path.basename(source_path)
        dest = os.path.join(DIR_CONOCIMIENTO, name)

        try:
            if os.path.abspath(source_path) != os.path.abspath(dest):
                shutil.copy2(source_path, dest)

            if ext == ".pdf" and PdfReader:
                reader = PdfReader(dest)
                for page in reader.pages[:10]:
                    text_data += (page.extract_text() or "") + "\n"
            elif ext == ".docx" and docx:
                doc = docx.Document(dest)
                text_data = "\n".join([p.text for p in doc.paragraphs if p.text])
            elif ext == ".txt":
                with open(dest, "r", encoding="utf-8") as f:
                    text_data = f.read()

            if text_data.strip():
                snippet = f"\n\n--- {name} ---\n{text_data[:3000]}"
                self.txt_biblio.insert(tk.END, snippet)
                messagebox.showinfo("Guardado", f"'{name}' agregado al contexto.")
        except Exception as e:
            messagebox.showerror("Error", f"Fallo al leer archivo: {e}")

    def salvar_perfil(self) -> None:
        self.config["perfil_nombre"] = self.ent_nombre.get().strip() or "Nombre"
        self.config["perfil_rol"] = self.ent_rol.get().strip()
        self.config["perfil_empresa"] = self.ent_empresa.get().strip() or "Teinek Solutions"
        self.config["perfil_bibliografia"] = self.txt_biblio.get("1.0", tk.END).strip()
        guardar_configuracion(self.config)
        messagebox.showinfo("Éxito", "Perfil guardado correctamente.")

    def salvar_credenciales(self) -> None:
        try:
            self.config["timeout_wa_web"] = max(10, int(self.spin_wa_timeout.get()))
        except ValueError:
            self.config["timeout_wa_web"] = 18

        self.config["llm_provider"] = self.combo_prov.get()
        self.config["llm_key"] = self.ent_key.get().strip()
        self.config["llm_model"] = MODELOS_IA.get(self.combo_prov.get(), "openai/gpt-oss-120b")
        self.config["green_id"] = self.ent_gid.get().strip()
        self.config["green_token"] = self.ent_gtok.get().strip()
        guardar_configuracion(self.config)
        messagebox.showinfo("Éxito", "Ajustes guardados.")

    def recargar_tabla_historial(self) -> None:
        for r in self.tree.get_children():
            self.tree.delete(r)
        for d in obtener_historial():
            self.tree.insert("", tk.END, values=(
                d.get("canal", "WhatsApp Web"),
                d.get("destinatario"),
                d.get("mensaje"),
                d.get("adjunto", "No"),
                d.get("fecha_envio"),
                d.get("estado"),
            ))

# ==============================================================================
# 8. PUNTO DE ENTRADA PRINCIPAL
# ==============================================================================
if __name__ == "__main__":
    if DND_AVAILABLE and TkinterDnD is not None:
        app_root = TkinterDnD.Tk()
    else:
        app_root = tk.Tk()

    app_instance = WaitSenderApp(app_root)
    app_root.protocol(
        "WM_DELETE_WINDOW",
        lambda: (app_instance.scheduler.shutdown(wait=False), app_root.destroy())
    )
    app_root.mainloop()
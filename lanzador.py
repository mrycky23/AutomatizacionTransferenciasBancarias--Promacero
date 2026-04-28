"""
=============================================================
  LANZADOR DEL SERVIDOR - Firesoft Transferencias
=============================================================
  Este script:
  1. Inicia api_servidor.py en segundo plano
  2. Inicia cloudflared para el túnel
  3. Muestra ícono en la bandeja del sistema (system tray)
  4. Muestra la URL pública para copiar en los celulares
  5. Se puede cerrar desde la bandeja sin abrir consola

  Instalar dependencias:
    pip install pystray pillow

  Para compilar como .exe:
    pyinstaller --onefile --windowed --name "Firesoft-Servidor" lanzador.py
=============================================================
"""

import subprocess
import threading
import sys
import os
import time
import re
import json
import webbrowser
from datetime import datetime

import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

try:
    import pystray
    from pystray import MenuItem as item
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

# ==============================
# CONFIGURACIÓN
# ==============================

CARPETA_BASE     = os.path.dirname(os.path.abspath(__file__))
SERVIDOR_SCRIPT  = os.path.join(CARPETA_BASE, "api_servidor.py")
CLOUDFLARED_EXE  = os.path.join(CARPETA_BASE, "cloudflared.exe")
PUERTO           = 8080
APP_URL          = "https://promacero.com"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ==============================
# ESTADO GLOBAL
# ==============================

proceso_servidor   = None
proceso_cloudflare = None
url_publica        = ""
estado             = "detenido"  # detenido | iniciando | activo | error
log_lines          = []

# ==============================
# CREAR ÍCONO PARA BANDEJA
# ==============================

def crear_icono_imagen(color="#4f8ef7"):
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse([4, 4, 60, 60], fill=color)
    d.text((20, 20), "FS", fill="white")
    return img


# ==============================
# FUNCIONES SERVIDOR
# ==============================

def agregar_log(msg):
    global log_lines
    timestamp = datetime.now().strftime("%H:%M:%S")
    linea = f"[{timestamp}] {msg}"
    log_lines.append(linea)
    if len(log_lines) > 200:
        log_lines = log_lines[-200:]
    if ventana and ventana.winfo_exists():
        try:
            ventana.after(0, lambda: actualizar_log_ui(linea))
        except:
            pass
    print(linea)


def iniciar_servidor():
    global proceso_servidor, estado

    if not os.path.exists(SERVIDOR_SCRIPT):
        agregar_log(f"ERROR: No se encontró {SERVIDOR_SCRIPT}")
        estado = "error"
        return

    agregar_log("Iniciando servidor Flask...")
    try:
        proceso_servidor = subprocess.Popen(
            [sys.executable, SERVIDOR_SCRIPT],
            cwd=CARPETA_BASE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        def leer_output():
            for linea in proceso_servidor.stdout:
                txt = linea.decode("utf-8", errors="ignore").strip()
                if txt:
                    agregar_log(f"Flask: {txt}")

        threading.Thread(target=leer_output, daemon=True).start()
        time.sleep(2)
        agregar_log(f"Servidor Flask activo en puerto {PUERTO} ✅")

    except Exception as e:
        agregar_log(f"ERROR servidor: {e}")
        estado = "error"


def iniciar_cloudflare():
    global proceso_cloudflare, url_publica, estado

    if not os.path.exists(CLOUDFLARED_EXE):
        agregar_log("cloudflared.exe no encontrado — operando solo en red local")
        url_publica = f"http://localhost:{PUERTO}"
        estado = "activo"
        actualizar_url_ui()
        return

    agregar_log("Iniciando túnel Cloudflare...")
    try:
        proceso_cloudflare = subprocess.Popen(
            [CLOUDFLARED_EXE, "tunnel", "--url", f"http://localhost:{PUERTO}"],
            cwd=CARPETA_BASE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        def leer_url():
            global url_publica, estado
            for linea in proceso_cloudflare.stdout:
                txt = linea.decode("utf-8", errors="ignore").strip()
                if txt:
                    agregar_log(f"Cloudflare: {txt}")

                match = re.search(r"https://[a-z0-9\-]+\.trycloudflare\.com", txt)
                if match:
                    url_publica = match.group(0)
                    estado = "activo"
                    agregar_log(f"URL pública: {url_publica} ✅")
                    if ventana and ventana.winfo_exists():
                        ventana.after(0, actualizar_url_ui)
                    break

            for linea in proceso_cloudflare.stdout:
                txt = linea.decode("utf-8", errors="ignore").strip()
                if txt:
                    agregar_log(f"Cloudflare: {txt}")

        threading.Thread(target=leer_url, daemon=True).start()

    except Exception as e:
        agregar_log(f"ERROR cloudflare: {e}")
        url_publica = f"http://localhost:{PUERTO}"
        estado = "activo"
        actualizar_url_ui()


def iniciar_todo():
    global estado
    estado = "iniciando"
    actualizar_estado_ui()

    def worker():
        iniciar_servidor()
        iniciar_cloudflare()

    threading.Thread(target=worker, daemon=True).start()


def detener_todo():
    global proceso_servidor, proceso_cloudflare, estado
    agregar_log("Deteniendo servicios...")

    if proceso_cloudflare:
        try:
            proceso_cloudflare.terminate()
        except:
            pass
        proceso_cloudflare = None

    if proceso_servidor:
        try:
            proceso_servidor.terminate()
        except:
            pass
        proceso_servidor = None

    estado = "detenido"
    agregar_log("Servicios detenidos.")
    if ventana and ventana.winfo_exists():
        ventana.after(0, actualizar_estado_ui)


# ==============================
# INTERFAZ GRÁFICA
# ==============================

ventana = None
lbl_estado    = None
lbl_url       = None
txt_log       = None
btn_iniciar   = None
btn_detener   = None
lbl_url_valor = None


def actualizar_log_ui(linea):
    if txt_log:
        txt_log.configure(state="normal")
        txt_log.insert("end", linea + "\n")
        txt_log.see("end")
        txt_log.configure(state="disabled")


def actualizar_url_ui():
    if lbl_url_valor:
        lbl_url_valor.configure(text=url_publica or "Obteniendo URL...")
    if lbl_estado:
        color = "#22c55e" if estado == "activo" else "#f59e0b" if estado == "iniciando" else "#ef4444"
        texto = "● Activo" if estado == "activo" else "● Iniciando..." if estado == "iniciando" else "● Detenido"
        lbl_estado.configure(text=texto, text_color=color)


def actualizar_estado_ui():
    actualizar_url_ui()
    if btn_iniciar and btn_detener:
        if estado in ("activo", "iniciando"):
            btn_iniciar.configure(state="disabled")
            btn_detener.configure(state="normal")
        else:
            btn_iniciar.configure(state="normal")
            btn_detener.configure(state="disabled")


def copiar_url():
    if url_publica:
        ventana.clipboard_clear()
        ventana.clipboard_append(url_publica)
        messagebox.showinfo("Copiado", f"URL copiada:\n{url_publica}")


def abrir_app():
    webbrowser.open(APP_URL)


def crear_ventana():
    global ventana, lbl_estado, lbl_url_valor, txt_log, btn_iniciar, btn_detener

    ventana = ctk.CTk()
    ventana.title("Firesoft · Servidor Transferencias")
    ventana.geometry("540x580")
    ventana.resizable(False, False)

    ventana.protocol("WM_DELETE_WINDOW", minimizar_a_bandeja)

    # Header
    header = ctk.CTkFrame(ventana, fg_color="#1a1a2e", corner_radius=0)
    header.pack(fill="x")

    ctk.CTkLabel(
        header,
        text="⚡ Firesoft · Servidor de Transferencias",
        font=ctk.CTkFont(size=16, weight="bold")
    ).pack(pady=16)

    # Estado
    frame_estado = ctk.CTkFrame(ventana)
    frame_estado.pack(fill="x", padx=20, pady=(10, 4))

    ctk.CTkLabel(frame_estado, text="Estado:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=12, pady=10)
    lbl_estado = ctk.CTkLabel(frame_estado, text="● Detenido", text_color="#ef4444", font=ctk.CTkFont(size=14))
    lbl_estado.pack(side="left", pady=10)

    # URL pública
    frame_url = ctk.CTkFrame(ventana)
    frame_url.pack(fill="x", padx=20, pady=4)

    ctk.CTkLabel(frame_url, text="URL pública:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=12, pady=10)
    lbl_url_valor = ctk.CTkLabel(frame_url, text="—", text_color="#4f8ef7", font=ctk.CTkFont(size=12), wraplength=300)
    lbl_url_valor.pack(side="left", pady=10, padx=4)

    ctk.CTkButton(
        frame_url, text="📋 Copiar",
        width=80, height=28,
        command=copiar_url
    ).pack(side="right", padx=12, pady=8)

    # Botones
    frame_btns = ctk.CTkFrame(ventana, fg_color="transparent")
    frame_btns.pack(fill="x", padx=20, pady=8)

    btn_iniciar = ctk.CTkButton(
        frame_btns,
        text="▶ Iniciar servidor",
        fg_color="#22c55e",
        hover_color="#16a34a",
        height=38,
        command=lambda: [iniciar_todo(), actualizar_estado_ui()]
    )
    btn_iniciar.pack(side="left", padx=4, expand=True, fill="x")

    btn_detener = ctk.CTkButton(
        frame_btns,
        text="■ Detener",
        fg_color="#ef4444",
        hover_color="#b91c1c",
        height=38,
        state="disabled",
        command=detener_todo
    )
    btn_detener.pack(side="left", padx=4, expand=True, fill="x")

    ctk.CTkButton(
        frame_btns,
        text="🌐 Abrir app",
        fg_color="#3b82f6",
        hover_color="#1d4ed8",
        height=38,
        width=110,
        command=abrir_app
    ).pack(side="left", padx=4)

    # Log
    ctk.CTkLabel(ventana, text="Registro de actividad:", anchor="w").pack(fill="x", padx=20, pady=(8, 2))

    txt_log = ctk.CTkTextbox(ventana, height=260, state="disabled", font=ctk.CTkFont(family="Courier", size=11))
    txt_log.pack(fill="both", expand=True, padx=20, pady=(0, 16))

    # Nota de bandeja
    ctk.CTkLabel(
        ventana,
        text="Al cerrar la ventana el servidor sigue activo en la bandeja del sistema",
        text_color="gray",
        font=ctk.CTkFont(size=11)
    ).pack(pady=(0, 12))

    return ventana


# ==============================
# BANDEJA DEL SISTEMA
# ==============================

icono_bandeja = None


def minimizar_a_bandeja():
    if TRAY_AVAILABLE:
        ventana.withdraw()
        iniciar_bandeja()
    else:
        if messagebox.askyesno("Salir", "¿Deseas detener el servidor y salir?"):
            detener_todo()
            ventana.destroy()


def mostrar_ventana(icon=None, item=None):
    if icono_bandeja:
        icono_bandeja.stop()
    ventana.after(0, ventana.deiconify)


def salir_completo(icon=None, item=None):
    detener_todo()
    if icono_bandeja:
        icono_bandeja.stop()
    ventana.after(0, ventana.destroy)


def iniciar_bandeja():
    global icono_bandeja

    img = crear_icono_imagen("#22c55e" if estado == "activo" else "#ef4444")

    menu = pystray.Menu(
        item("Mostrar ventana", mostrar_ventana, default=True),
        item("Copiar URL pública", lambda i, it: copiar_url()),
        pystray.Menu.SEPARATOR,
        item("Detener y salir", salir_completo)
    )

    icono_bandeja = pystray.Icon("firesoft", img, "Firesoft Servidor", menu)

    threading.Thread(target=icono_bandeja.run, daemon=True).start()


# ==============================
# INICIO AUTOMÁTICO CON WINDOWS
# ==============================

def configurar_inicio_automatico():
    """
    Agrega el lanzador al registro de Windows para que
    inicie automáticamente con el sistema.
    """
    import winreg

    ruta_exe = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
    clave = r"Software\Microsoft\Windows\CurrentVersion\Run"

    try:
        reg = winreg.OpenKey(winreg.HKEY_CURRENT_USER, clave, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(reg, "FiresoftServidor", 0, winreg.REG_SZ, f'"{ruta_exe}"')
        winreg.CloseKey(reg)
        agregar_log("Inicio automático configurado ✅")
        messagebox.showinfo("Inicio automático", "✅ El servidor iniciará automáticamente con Windows.")
    except Exception as e:
        agregar_log(f"Error configurando inicio automático: {e}")
        messagebox.showerror("Error", str(e))


def quitar_inicio_automatico():
    import winreg
    clave = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        reg = winreg.OpenKey(winreg.HKEY_CURRENT_USER, clave, 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(reg, "FiresoftServidor")
        winreg.CloseKey(reg)
        messagebox.showinfo("Inicio automático", "El inicio automático fue desactivado.")
    except:
        messagebox.showinfo("Info", "No estaba configurado el inicio automático.")


# ==============================
# MAIN
# ==============================

if __name__ == "__main__":
    v = crear_ventana()

    # Menú extra en la barra
    menu_bar = tk.Menu(v)
    menu_config = tk.Menu(menu_bar, tearoff=0)
    menu_config.add_command(label="✅ Activar inicio con Windows", command=configurar_inicio_automatico)
    menu_config.add_command(label="❌ Desactivar inicio con Windows", command=quitar_inicio_automatico)
    menu_bar.add_cascade(label="Configuración", menu=menu_config)
    v.configure(menu=menu_bar)

    # Auto-iniciar al abrir
    v.after(1000, lambda: [iniciar_todo(), actualizar_estado_ui()])

    v.mainloop()
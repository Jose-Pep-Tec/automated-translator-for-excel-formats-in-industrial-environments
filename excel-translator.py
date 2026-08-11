"""
Traductor de archivos Excel (español → inglés técnico)
Preserva formatos, colores, fórmulas y estilos.
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.cell.cell import MergedCell
from deep_translator import GoogleTranslator
from deep_translator.exceptions import RequestError, TooManyRequests
from pathlib import Path
import threading
import time
import traceback
import json
import re

# Configuración de la interfaz gráfica
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# Parámetros de traducción
BATCH_SIZE = 40  # Textos por petición (límite seguro para Google)
PAUSE_BETWEEN = 0.6  # Segundos entre lotes
MAX_RETRIES = 4  # Reintentos antes de rendirse
BACKOFF_BASE = 2.0  # Segundos base para espera exponencial
MAX_TEXT_LEN = 4800  # Google Translate corta a ~5000 caracteres
DICCIONARIO_FILE = "custom_dictionary.json"  # Archivo de diccionario personalizado


# Clase para manejar el diccionario personalizado
class CustomDictionary:
    def __init__(self):
        self.diccionario = {}
        self.cargar_diccionario()
    
    def cargar_diccionario(self):
        """Carga el diccionario desde archivo JSON"""
        try:
            if Path(DICCIONARIO_FILE).exists():
                with open(DICCIONARIO_FILE, 'r', encoding='utf-8') as f:
                    self.diccionario = json.load(f)
        except Exception as e:
            print(f"Error cargando diccionario: {e}")
            self.diccionario = {}
    
    def guardar_diccionario(self):
        """Guarda el diccionario en archivo JSON"""
        try:
            with open(DICCIONARIO_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.diccionario, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error guardando diccionario: {e}")
    
    def agregar_palabra(self, original, traduccion):
        """Agrega o actualiza una palabra en el diccionario"""
        self.diccionario[original.lower()] = traduccion
        self.guardar_diccionario()
    
    def eliminar_palabra(self, original):
        """Elimina una palabra del diccionario"""
        if original.lower() in self.diccionario:
            del self.diccionario[original.lower()]
            self.guardar_diccionario()
            return True
        return False
    
    def traducir_con_diccionario(self, texto):
        """Aplica el diccionario personalizado a un texto"""
        if not texto or not isinstance(texto, str):
            return texto
        
        texto_traducido = texto
        # Ordenar por longitud (más largas primero) para evitar reemplazos parciales
        palabras_ordenadas = sorted(self.diccionario.keys(), key=len, reverse=True)
        
        for palabra_original in palabras_ordenadas:
            # Buscar la palabra como palabra completa (con delimitadores)
            patron = r'\b' + re.escape(palabra_original) + r'\b'
            texto_traducido = re.sub(patron, self.diccionario[palabra_original], texto_traducido, flags=re.IGNORECASE)
        
        return texto_traducido


# Funciones auxiliares
def es_celda_escribible(cell) -> bool:
    """Verifica si se puede leer y escribir la celda de forma segura"""
    return not isinstance(cell, MergedCell)


def traducir_lote(translator: GoogleTranslator, textos: list[str], custom_dict: CustomDictionary = None, usar_formato_dual: bool = True, para_nombre_hoja: bool = False) -> list[str]:
    """
    Traduce una lista de textos con reintentos y espera progresiva.
    - usar_formato_dual: mantiene el texto original + " / " + traducción
    - para_nombre_hoja: no aplica formato dual y limpia caracteres inválidos
    """
    # Aplicar diccionario personalizado primero
    textos_procesados = textos.copy()
    if custom_dict:
        for i, texto in enumerate(textos_procesados):
            if texto:
                textos_procesados[i] = custom_dict.traducir_con_diccionario(texto)
    
    for intento in range(MAX_RETRIES):
        try:
            resultado = translator.translate_batch(textos_procesados)
            resultado = [r if r is not None else t for r, t in zip(resultado, textos_procesados)]
            
            # Caso: nombre de hoja - limpiar caracteres inválidos
            if para_nombre_hoja:
                for i, traducido in enumerate(resultado):
                    for ch in r'\/?*[]:':
                        traducido = traducido.replace(ch, '_')
                    resultado[i] = traducido[:31]  # Excel permite máximo 31 caracteres
            
            # Caso: modo dual para celdas
            elif usar_formato_dual:
                for i, (original, traducido) in enumerate(zip(textos, resultado)):
                    if original and traducido and original != traducido:
                        if not (" / " in str(original) and str(traducido) in str(original)):
                            resultado[i] = f"{original} / {traducido}"
                    else:
                        resultado[i] = original
            
            return resultado
            
        except TooManyRequests:
            espera = BACKOFF_BASE ** (intento + 1)
            time.sleep(espera)
        except RequestError as e:
            espera = BACKOFF_BASE ** (intento + 1)
            time.sleep(espera)
        except Exception:
            time.sleep(BACKOFF_BASE)
    
    raise RuntimeError(f"Lote falló tras {MAX_RETRIES} intentos")


def truncar_si_necesario(texto: str) -> tuple[str, bool]:
    """Trunca texto si supera el límite permitido por Google Translate"""
    if len(texto) > MAX_TEXT_LEN:
        return texto[:MAX_TEXT_LEN], True
    return texto, False


# Aplicación principal
class ExcelTranslator(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Traductor de Excel con Diccionario Personalizado")
        self.geometry("900x800")
        self.resizable(True, True)
        self.custom_dict = CustomDictionary()
        self._build_ui()
    
    def _build_ui(self):
        # Título
        ctk.CTkLabel(
            self,
            text="Traductor de Excel",
            font=("Courier New", 22, "bold"),
            text_color="#39ff14",
        ).pack(pady=(22, 4))

        ctk.CTkLabel(
            self,
            text="Traduce archivos .xlsx preservando formatos, fórmulas y estilos",
            font=("Courier New", 11),
            text_color="#888888",
        ).pack(pady=(0, 14))
        
        # Panel de pestañas
        self.tabview = ctk.CTkTabview(self, width=850, height=650)
        self.tabview.pack(padx=20, pady=10, fill="both", expand=True)
        
        tab_principal = self.tabview.add("Traductor")
        tab_dict = self.tabview.add("Diccionario")
        
        # ── Pestaña principal ─────────────────────────────────────────────────
        
        # Selector de carpeta
        frame_path = ctk.CTkFrame(tab_principal, corner_radius=10)
        frame_path.pack(fill="x", padx=24, pady=6)

        self.ruta_carpeta = ctk.StringVar()
        ctk.CTkEntry(
            frame_path,
            textvariable=self.ruta_carpeta,
            placeholder_text="Selecciona una carpeta con archivos .xlsx...",
            font=("Courier New", 12),
            width=540,
        ).pack(side="left", padx=12, pady=10, expand=True, fill="x")

        ctk.CTkButton(
            frame_path,
            text="Carpeta",
            width=110,
            command=self._seleccionar,
            fg_color="#1a3a1a",
            hover_color="#2a5a2a",
            font=("Courier New", 12, "bold"),
        ).pack(side="right", padx=12)

        # Identificador de salida
        frame_id = ctk.CTkFrame(tab_principal, corner_radius=10)
        frame_id.pack(fill="x", padx=24, pady=(6, 2))

        ctk.CTkLabel(
            frame_id,
            text="Prefijo para archivos de salida:",
            font=("Courier New", 11),
            text_color="#aaaaaa",
        ).pack(side="left", padx=(12, 6), pady=8)

        self.var_identificador = ctk.StringVar(value="")
        self.entry_identificador = ctk.CTkEntry(
            frame_id,
            textvariable=self.var_identificador,
            width=100,
            font=("Courier New", 12),
            placeholder_text="ej: EN",
            justify="center",
        )
        self.entry_identificador.pack(side="left", padx=(0, 10), pady=8)

        ctk.CTkLabel(
            frame_id,
            text="← (dejar vacío para no agregar prefijo)",
            font=("Courier New", 10),
            text_color="#666666",
        ).pack(side="left", padx=4)
        
        # Formato dual
        self.var_formato_dual = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            frame_id,
            text="Formato dual (original / traducción)",
            variable=self.var_formato_dual,
            font=("Courier New", 11),
        ).pack(side="right", padx=20)

        # Opciones
        frame_opts = ctk.CTkFrame(tab_principal, corner_radius=10, fg_color="transparent")
        frame_opts.pack(fill="x", padx=24, pady=2)

        self.var_nombres_hoja = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            frame_opts,
            text="Traducir nombres de hoja",
            variable=self.var_nombres_hoja,
            font=("Courier New", 11),
        ).pack(side="left", padx=6)

        self.var_traducir_nombre_archivo = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            frame_opts,
            text="Traducir nombre del archivo",
            variable=self.var_traducir_nombre_archivo,
            font=("Courier New", 11),
        ).pack(side="left", padx=20)

        self.var_sobrescribir = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            frame_opts,
            text="Sobrescribir archivos existentes",
            variable=self.var_sobrescribir,
            font=("Courier New", 11),
        ).pack(side="left", padx=20)

        # Área de registro
        self.txt_log = ctk.CTkTextbox(
            tab_principal,
            height=200,
            font=("Courier New", 11),
            text_color="#cccccc",
            fg_color="#0d1a0d",
            corner_radius=8,
        )
        self.txt_log.pack(fill="both", padx=24, pady=10, expand=True)

        # Barra de progreso
        frame_prog = ctk.CTkFrame(tab_principal, fg_color="transparent")
        frame_prog.pack(fill="x", padx=24)

        self.lbl_progress = ctk.CTkLabel(
            frame_prog,
            text="En espera...",
            font=("Courier New", 10),
            text_color="#666666",
        )
        self.lbl_progress.pack(anchor="w")

        self.progress = ctk.CTkProgressBar(tab_principal, height=14, corner_radius=6)
        self.progress.pack(fill="x", padx=24, pady=(2, 10))
        self.progress.set(0)

        # Botón principal
        self.btn_run = ctk.CTkButton(
            tab_principal,
            text="TRADUCIR",
            command=self._iniciar,
            fg_color="#1e5128",
            hover_color="#2d7a3a",
            font=("Courier New", 15, "bold"),
            height=44,
            corner_radius=10,
        )
        self.btn_run.pack(pady=(4, 20))
        
        # ── Pestaña de diccionario ───────────────────────────────────────────
        self._build_dictionary_tab(tab_dict)
    
    def _build_dictionary_tab(self, parent):
        """Construye la interfaz del diccionario personalizado"""
        
        ctk.CTkLabel(
            parent,
            text="Diccionario Personalizado",
            font=("Courier New", 16, "bold"),
            text_color="#39ff14",
        ).pack(pady=(15, 5))
        
        ctk.CTkLabel(
            parent,
            text="Estas traducciones se aplican antes de usar Google Translate",
            font=("Courier New", 11),
            text_color="#888888",
        ).pack(pady=(0, 15))
        
        # Frame para agregar entradas
        frame_edit = ctk.CTkFrame(parent, corner_radius=10)
        frame_edit.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(frame_edit, text="Palabra original:", font=("Courier New", 11)).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.entry_original = ctk.CTkEntry(frame_edit, width=200, font=("Courier New", 11))
        self.entry_original.grid(row=0, column=1, padx=10, pady=10)
        
        ctk.CTkLabel(frame_edit, text="Traducción deseada:", font=("Courier New", 11)).grid(row=0, column=2, padx=10, pady=10, sticky="w")
        self.entry_traduccion = ctk.CTkEntry(frame_edit, width=200, font=("Courier New", 11))
        self.entry_traduccion.grid(row=0, column=3, padx=10, pady=10)
        
        ctk.CTkButton(
            frame_edit,
            text="Agregar",
            command=self._agregar_palabra_dict,
            fg_color="#1e5128",
            hover_color="#2d7a3a",
            width=100
        ).grid(row=0, column=4, padx=10, pady=10)
        
        # Lista de palabras
        frame_lista = ctk.CTkFrame(parent, corner_radius=10)
        frame_lista.pack(fill="both", expand=True, padx=20, pady=10)
        
        from tkinter import ttk
        columns = ("original", "traduccion")
        self.tree_dict = ttk.Treeview(frame_lista, columns=columns, show="headings", height=15)
        self.tree_dict.heading("original", text="Palabra Original")
        self.tree_dict.heading("traduccion", text="Traducción Personalizada")
        self.tree_dict.column("original", width=200)
        self.tree_dict.column("traduccion", width=200)
        
        scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=self.tree_dict.yview)
        self.tree_dict.configure(yscrollcommand=scrollbar.set)
        
        self.tree_dict.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
        
        ctk.CTkButton(
            parent,
            text="Eliminar seleccionados",
            command=self._eliminar_palabras_dict,
            fg_color="#8b0000",
            hover_color="#a00000"
        ).pack(pady=10)
        
        self._actualizar_lista_dict()
    
    def _agregar_palabra_dict(self):
        original = self.entry_original.get().strip()
        traduccion = self.entry_traduccion.get().strip()
        
        if not original or not traduccion:
            messagebox.showwarning("Campos vacíos", "Por favor completa ambos campos")
            return
        
        self.custom_dict.agregar_palabra(original, traduccion)
        self._actualizar_lista_dict()
        self.entry_original.delete(0, "end")
        self.entry_traduccion.delete(0, "end")
        self.log(f"Palabra agregada al diccionario: {original} → {traduccion}", "OK")
    
    def _eliminar_palabras_dict(self):
        seleccion = self.tree_dict.selection()
        if not seleccion:
            messagebox.showwarning("Sin selección", "Selecciona al menos una palabra para eliminar")
            return
        
        for item in seleccion:
            original = self.tree_dict.item(item)["values"][0]
            self.custom_dict.eliminar_palabra(original)
        
        self._actualizar_lista_dict()
        self.log(f"{len(seleccion)} palabra(s) eliminada(s) del diccionario", "OK")
    
    def _actualizar_lista_dict(self):
        for item in self.tree_dict.get_children():
            self.tree_dict.delete(item)
        
        for original, traduccion in self.custom_dict.diccionario.items():
            self.tree_dict.insert("", "end", values=(original, traduccion))

    def log(self, mensaje: str, nivel: str = "INFO"):
        iconos = {"INFO": "·", "OK": "✓", "WARN": "!", "ERR": "✗", "HEAD": "—"}
        icono = iconos.get(nivel, "·")
        linea = f"[{time.strftime('%H:%M:%S')}] {icono} {mensaje}\n"
        self.txt_log.insert("end", linea)
        self.txt_log.see("end")

    def set_progress(self, valor: float, texto: str = ""):
        self.progress.set(max(0.0, min(1.0, valor)))
        if texto:
            self.lbl_progress.configure(text=texto)

    def _seleccionar(self):
        path = filedialog.askdirectory()
        if path:
            self.ruta_carpeta.set(path)

    def _iniciar(self):
        if not self.ruta_carpeta.get():
            messagebox.showwarning("Sin carpeta", "Selecciona una carpeta primero.")
            return
        self.txt_log.delete("1.0", "end")
        self.btn_run.configure(state="disabled")
        threading.Thread(target=self._proceso, daemon=True).start()

    def _proceso(self):
        folder = Path(self.ruta_carpeta.get())
        sobrescribir = self.var_sobrescribir.get()
        traducir_hojas = self.var_nombres_hoja.get()
        traducir_nombre_archivo = self.var_traducir_nombre_archivo.get()
        identificador = self.var_identificador.get().strip()
        usar_formato_dual = self.var_formato_dual.get()

        # Buscar archivos .xlsx
        todos = list(folder.glob("*.xlsx"))
        
        # Filtrar según sobrescribir
        if not sobrescribir and identificador:
            archivos = [
                f for f in todos
                if not f.stem.startswith(f"{identificador}_")
            ]
        else:
            archivos = todos

        if not archivos:
            self.log("No se encontraron archivos .xlsx para procesar.", "WARN")
            self.btn_run.configure(state="normal")
            return

        self.log(f"Archivos encontrados: {len(archivos)}", "HEAD")
        self.log(f"Modo de traducción: {'Dual (original / traducción)' if usar_formato_dual else 'Solo traducción'}", "INFO")
        self.log(f"Diccionario personalizado: {len(self.custom_dict.diccionario)} palabras cargadas", "INFO")
        self.log(f"Traducir nombres de hoja: {'Sí' if traducir_hojas else 'No'}", "INFO")
        self.log(f"Traducir nombre del archivo: {'Sí' if traducir_nombre_archivo else 'No'}", "INFO")
        self.log(f"Prefijo de salida: '{identificador}' {'(se agregará al inicio)' if identificador else '(no se agregará prefijo)'}", "INFO")
        self.log(f"Sobrescribir: {'Sí' if sobrescribir else 'No'}", "INFO")

        translator = GoogleTranslator(source="auto", target="en")

        stats = {"ok": 0, "fail": 0, "celdas_ok": 0, "celdas_warn": 0}

        for file_idx, file_path in enumerate(archivos):
            archivo_ok = self._traducir_archivo(
                file_path, translator, traducir_hojas,
                traducir_nombre_archivo, identificador,
                file_idx, len(archivos), stats, usar_formato_dual
            )
            if archivo_ok:
                stats["ok"] += 1
            else:
                stats["fail"] += 1

        self.log("", "HEAD")
        self.log(f"FINALIZADO — Archivos OK: {stats['ok']} | Fallidos: {stats['fail']}", "HEAD")
        self.log(f"Celdas traducidas: {stats['celdas_ok']} | Con advertencia: {stats['celdas_warn']}", "HEAD")
        self.set_progress(1.0, "Completado")
        self.btn_run.configure(state="normal")
        messagebox.showinfo(
            "Traducción completa",
            f"Archivos procesados: {stats['ok']}\n"
            f"Fallidos: {stats['fail']}\n"
            f"Celdas traducidas: {stats['celdas_ok']}\n"
            f"Celdas con advertencia: {stats['celdas_warn']}"
        )

    def _traducir_archivo(
        self, file_path: Path, translator, traducir_hojas: bool,
        traducir_nombre_archivo: bool, identificador: str,
        file_idx: int, total_files: int, stats: dict, usar_formato_dual: bool
    ) -> bool:
        
        self.log(f"{'─'*55}", "HEAD")
        self.log(f"[{file_idx+1}/{total_files}] {file_path.name}")

        # Cargar el libro de trabajo
        try:
            wb = openpyxl.load_workbook(file_path, data_only=False)
        except PermissionError:
            self.log(f"Archivo bloqueado (¿abierto en Excel?): {file_path.name}", "ERR")
            return False
        except Exception as e:
            self.log(f"No se pudo abrir {file_path.name}: {e}", "ERR")
            return False

        # Procesar hojas
        for sheet in wb.worksheets:
            nombre_original = sheet.title
            self.log(f"  Hoja: «{nombre_original}»")

            # Recolectar celdas traducibles
            celdas_a_traducir = []

            for row in sheet.iter_rows():
                for cell in row:
                    if not es_celda_escribible(cell):
                        continue
                    val = cell.value
                    if not isinstance(val, str):
                        continue
                    if val.startswith("="):
                        continue
                    texto = val.strip()
                    if not texto:
                        continue
                    celdas_a_traducir.append((cell, texto))

            if celdas_a_traducir:
                total_celdas = len(celdas_a_traducir)
                self.log(f"    {total_celdas} celdas a traducir...")

                for lote_inicio in range(0, total_celdas, BATCH_SIZE):
                    lote_refs = celdas_a_traducir[lote_inicio: lote_inicio + BATCH_SIZE]
                    lote_textos = []
                    truncadas = []

                    for cell, texto in lote_refs:
                        texto_seg, fue_truncado = truncar_si_necesario(texto)
                        lote_textos.append(texto_seg)
                        if fue_truncado:
                            truncadas.append(cell.coordinate)

                    if truncadas:
                        self.log(f"    Texto truncado en: {', '.join(truncadas)}", "WARN")
                        stats["celdas_warn"] += len(truncadas)

                    try:
                        traducidos = traducir_lote(translator, lote_textos, self.custom_dict, usar_formato_dual, para_nombre_hoja=False)
                    except RuntimeError as e:
                        self.log(f"    Lote falló: {e}", "ERR")
                        stats["celdas_warn"] += len(lote_refs)
                        time.sleep(PAUSE_BETWEEN)
                        continue

                    for (cell, _), texto_traducido in zip(lote_refs, traducidos):
                        try:
                            cell.value = texto_traducido
                            stats["celdas_ok"] += 1
                        except Exception as write_err:
                            self.log(f"    No se pudo escribir en {cell.coordinate}: {write_err}", "WARN")
                            stats["celdas_warn"] += 1

                    time.sleep(PAUSE_BETWEEN)

                    progreso_global = (file_idx + (lote_inicio + BATCH_SIZE) / total_celdas) / total_files
                    self.set_progress(progreso_global, f"Archivo {file_idx+1}/{total_files} · celda {min(lote_inicio + BATCH_SIZE, total_celdas)}/{total_celdas}")
            else:
                self.log(f"    Sin texto para traducir.")

            # Traducir nombre de hoja (solo si está marcado)
            if traducir_hojas and nombre_original.strip():
                try:
                    nombre_en = traducir_lote(translator, [nombre_original], self.custom_dict, usar_formato_dual=False, para_nombre_hoja=True)[0]
                    nuevo_nombre = nombre_en[:31]
                    if nuevo_nombre != nombre_original:
                        sheet.title = nuevo_nombre
                        self.log(f"    Hoja renombrada: «{nombre_original}» → «{nuevo_nombre}»")
                except Exception as e:
                    self.log(f"    No se pudo traducir nombre de hoja: {e}", "WARN")

        # Construir nombre de salida
        nombre_salida = file_path.name
        
        if traducir_nombre_archivo:
            # Traducir el nombre del archivo (sin extensión)
            try:
                nombre_sin_ext = file_path.stem
                nombre_traducido = traducir_lote(translator, [nombre_sin_ext], self.custom_dict, usar_formato_dual=False, para_nombre_hoja=False)[0]
                # Limpiar caracteres inválidos para sistema de archivos
                for ch in r'\/:*?"<>|':
                    nombre_traducido = nombre_traducido.replace(ch, '_')
                nombre_salida = f"{nombre_traducido}{file_path.suffix}"
                self.log(f"  Nombre traducido: «{file_path.name}» → «{nombre_salida}»")
            except Exception as e:
                self.log(f"  No se pudo traducir el nombre del archivo: {e}", "WARN")
                nombre_salida = file_path.name
        
        # Agregar prefijo si se especificó
        if identificador:
            nombre_salida = f"{identificador}_{nombre_salida}"
            self.log(f"  Agregando prefijo: {identificador}_")

        ruta_salida = file_path.parent / nombre_salida

        try:
            wb.save(ruta_salida)
            self.log(f"  Guardado: {nombre_salida}", "OK")
            return True
        except Exception as e:
            self.log(f"  Error al guardar: {e}", "ERR")
            return False


if __name__ == "__main__":
    app = ExcelTranslator()
    app.mainloop()

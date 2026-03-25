"""
Interfaz gráfica de usuario para el convertidor DOCX a PDF
Versión mejorada con layout 16:9 y dos columnas
"""

print("=" * 60)
print("CARGANDO GUI.PY - VERSION 2024-01-27-FINAL")
print("=" * 60)


import tkinter as tk
from tkinter import scrolledtext, ttk
import os
from PIL import Image, ImageTk

try:
    from tkinterdnd2 import DND_FILES
    DRAG_DROP_DISPONIBLE = True
except ImportError:
    DRAG_DROP_DISPONIBLE = False
    print("tkinterdnd2 no disponible - drag & drop deshabilitado")

from src.config import (
    WINDOW_TITLE,
    COLOR_SUCCESS, COLOR_ERROR, COLOR_INFO, 
    COLOR_NEUTRAL, COLOR_DISABLED,
    COLOR_LOGO_BG, COLOR_LOGO_SUCCESS,
    PROGRESS_BAR_STYLE, PROGRESS_BAR_COLORS
)


class GUI:
    """Clase que gestiona toda la interfaz gráfica de la aplicación"""

    def __init__(self, root, controller):
        """
        Inicializa la interfaz gráfica
        
        Args:
            root: Ventana principal de Tkinter (ya creada)
            controller: Instancia del controlador que maneja la lógica
        """
        self.root = root
        self.controller = controller

        # Inicializar variables de preview ANTES de crear interfaz
        self.logo_photo = None
        self.canvas_image_id = None
        self.canvas_logo_preview = None

        # --- Variables para Checkboxes ---
        # Encabezado y Pie
        self.var_add_logo = tk.BooleanVar(value=True)
        self.var_add_folder_code = tk.BooleanVar(value=True)
        self.var_add_header_line = tk.BooleanVar(value=True)
        self.var_add_footer_line = tk.BooleanVar(value=True)
        self.var_add_author = tk.BooleanVar(value=True)
        self.var_add_page_number = tk.BooleanVar(value=True)

        # Copiar
        self.var_respect_structure = tk.BooleanVar(value=True)
        self.var_copy_attachments = tk.BooleanVar(value=True)
        self.var_save_modified_dest = tk.BooleanVar(value=True)
        self.var_copy_as_pdf = tk.BooleanVar(value=True)
        self.var_auto_rename = tk.BooleanVar(value=False) 

        # Extensiones
        self.var_process_docx = tk.BooleanVar(value=True)
        self.var_process_docm = tk.BooleanVar(value=False)

        # Configurar ventana
        self.root.title(WINDOW_TITLE)
        self.root.geometry("1280x720")
        self.root.resizable(False, False) 

        self._crear_interfaz()

    def _crear_interfaz(self):
        """Crea todos los elementos de la interfaz en layout de dos columnas"""
        frame_principal = tk.Frame(self.root, padx=8, pady=8)
        frame_principal.pack(fill=tk.BOTH, expand=True)

        # ====
        # COLUMNA IZQUIERDA: Configuración
        # ====
        frame_izquierda = tk.Frame(frame_principal)
        frame_izquierda.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        self._crear_seccion_logo(frame_izquierda)
        self._crear_seccion_autor(frame_izquierda)
        self._crear_seccion_opciones_detalladas(frame_izquierda)

        # ====
        # COLUMNA DERECHA: Carpetas, Destino y Progreso
        # ====
        frame_derecha = tk.Frame(frame_principal)
        frame_derecha.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._crear_seccion_carpetas(frame_derecha)
        self._crear_seccion_excepciones(frame_derecha)
        self._crear_seccion_destino(frame_derecha)
        self._crear_boton_empezar(frame_derecha)
        self._crear_seccion_progreso(frame_derecha)
        self._crear_seccion_log(frame_derecha)

    def _crear_seccion_logo(self, parent):
        """Crea la sección de selección de logo - versión compacta"""
        frame = tk.LabelFrame(
            parent, 
            text="🖼️ Logo del Encabezado", 
            font=("Arial", 9, "bold"), 
            padx=8, 
            pady=3
        )
        frame.pack(fill=tk.X, pady=(0, 3))

        # Canvas para preview - Dimensiones fijas
        CANVAS_W, CANVAS_H = 500, 70
        self.canvas_preview_width = CANVAS_W
        self.canvas_preview_height = CANVAS_H

        self.canvas_logo_preview = tk.Canvas(
            frame,
            width=CANVAS_W,
            height=CANVAS_H,
            bg=COLOR_LOGO_BG,
            highlightthickness=1,
            highlightbackground="#cccccc"
        )
        self.canvas_logo_preview.pack(pady=(0, 5), fill=tk.X)

        # Placeholder inicial
        self.canvas_logo_preview.create_text(
            CANVAS_W // 2, CANVAS_H // 2,
            text="Arrastra aquí el logo (PNG/JPG)",
            fill="#999999",
            font=("Arial", 10),
            tags="placeholder"
        )

        # Habilitar drag & drop
        if DRAG_DROP_DISPONIBLE:
            self.canvas_logo_preview.drop_target_register(DND_FILES)
            self.canvas_logo_preview.dnd_bind('<<Drop>>', self.controller.drop_logo)

        # Botón examinar
        tk.Button(
            frame, 
            text="📂  Examinar Logo", 
            command=self.controller.examinar_logo,
            bg="#2563eb",
            fg="white", 
            font=("Arial", 8, "bold"),
            relief=tk.FLAT,
            cursor="hand2",
            pady=4
        ).pack(fill=tk.X)

    def _crear_seccion_autor(self, parent):
        """Crea la sección de autor - versión compacta"""
        frame = tk.Frame(parent)
        frame.pack(fill=tk.X, pady=(0, 3))
        
        tk.Label(
            frame, 
            text="📝 Autor:", 
            font=("Arial", 9, "bold")
        ).pack(side=tk.LEFT)
        
        self.entry_autor = tk.Entry(frame, font=("Arial", 9))
        self.entry_autor.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        # Cargar valor por defecto
        autor_default = self._cargar_autor_desde_archivo()
        self.entry_autor.insert(0, autor_default)

    
    
    
    
    ######################################################
    ######################################################
    
    def solicitar_raiz_archivo(self, nombre_completo, nombre_sin_ext, nombre_carpeta=None):
        """
        Muestra un diálogo personalizado para ingresar la raíz del archivo.
        
        Args:
            nombre_completo (str): Nombre completo del archivo (con extensión)
            nombre_sin_ext (str): Nombre sin extensión (sugerencia por defecto)
            nombre_carpeta (str): Nombre de la carpeta contenedora (opcional)
        
        Returns:
            str: Raíz ingresada por el usuario o None si cancela
        """
        import tkinter as tk
        
        print(f"🔥🔥🔥 MÉTODO NUEVO DE DIALOGO - VERSION FINAL 🔥🔥🔥")
        
        # Variable para almacenar el resultado
        resultado = {'valor': None, 'cerrado': False}
        
        # Crear ventana modal personalizada
        dialogo = tk.Toplevel(self.root)
        dialogo.title("🔤 Definir Raíz del Archivo")
        dialogo.geometry("550x280")
        dialogo.transient(self.root)
        dialogo.grab_set()
        
        # Centrar el diálogo
        dialogo.update_idletasks()
        x = (dialogo.winfo_screenwidth() // 2) - (550 // 2)
        y = (dialogo.winfo_screenheight() // 2) - (280 // 2)
        dialogo.geometry(f"550x280+{x}+{y}")
        
        # Mensaje
        frame_msg = tk.Frame(dialogo)
        frame_msg.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Mostrar carpeta si está disponible
        if nombre_carpeta:
            tk.Label(
                frame_msg,
                text=f"📁 Carpeta: {nombre_carpeta}",
                font=("Arial", 9),
                fg="#666666"
            ).pack(anchor=tk.W, pady=(0, 5))
        
        tk.Label(
            frame_msg,
            text="No se detectó patrón automático en:",
            font=("Arial", 10)
        ).pack(anchor=tk.W)
        
        tk.Label(
            frame_msg,
            text=f"📄 {nombre_completo}",
            font=("Arial", 10, "bold"),
            fg="#0066cc"
        ).pack(anchor=tk.W, pady=(5, 15))
        
        tk.Label(
            frame_msg,
            text="Ingresa la RAÍZ del nombre del archivo\n(se le añadirá el código de carpeta al principio):",
            font=("Arial", 9),
            justify=tk.LEFT
        ).pack(anchor=tk.W, pady=(0, 10))
        
        # Campo de entrada
        entry_raiz = tk.Entry(frame_msg, font=("Arial", 11), width=50)
        entry_raiz.pack(fill=tk.X, pady=(0, 20))
        entry_raiz.insert(0, nombre_sin_ext)
        entry_raiz.select_range(0, tk.END)
        entry_raiz.focus_set()
        
        # Funciones de botones
        def aceptar():
            valor = entry_raiz.get().strip()
            if valor:
                resultado['valor'] = valor
            else:
                resultado['valor'] = None
            resultado['cerrado'] = True
            dialogo.destroy()
        
        def cancelar():
            resultado['valor'] = None
            resultado['cerrado'] = True
            dialogo.destroy()
        
        # Botones
        frame_botones = tk.Frame(dialogo)
        frame_botones.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        tk.Button(
            frame_botones,
            text="✓ Aceptar",
            command=aceptar,
            bg="#28a745",
            fg="white",
            font=("Arial", 10, "bold"),
            width=15
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Button(
            frame_botones,
            text="✗ Cancelar",
            command=cancelar,
            bg="#dc3545",
            fg="white",
            font=("Arial", 10, "bold"),
            width=15
        ).pack(side=tk.LEFT)
        
        # Permitir Enter para aceptar
        entry_raiz.bind('<Return>', lambda e: aceptar())
        
        # Permitir Escape para cancelar
        dialogo.bind('<Escape>', lambda e: cancelar())
        
        # Manejar cierre de ventana
        dialogo.protocol("WM_DELETE_WINDOW", cancelar)
        
        # ESPERA MANUAL - Procesar eventos hasta que se cierre
        while not resultado['cerrado']:
            try:
                self.root.update()
                dialogo.update()
            except:
                break
        
        return resultado['valor']
        
    
    
    
    
    
    
    
    
    
    
    
    

    def _crear_seccion_carpetas(self, parent):
        """Crea la sección de carpetas a procesar"""
        frame = tk.LabelFrame(
            parent, 
            text="📁 Carpetas a Procesar", 
            font=("Arial", 8, "bold"), 
            padx=4, 
            pady=3
        )
        frame.pack(fill=tk.X, pady=(0, 3))

        # Listbox
        f_list = tk.Frame(frame)
        f_list.pack(fill=tk.X)
        
        self.listbox_carpetas = tk.Listbox(f_list, height=3, font=("Arial", 8))
        self.listbox_carpetas.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Habilitar drag & drop
        if DRAG_DROP_DISPONIBLE:
            self.listbox_carpetas.drop_target_register(DND_FILES)
            self.listbox_carpetas.dnd_bind('<<Drop>>', self.controller.drop_carpeta)

        # Botones de gestión
        f_btns = tk.Frame(frame)
        f_btns.pack(fill=tk.X, pady=2)
        
        tk.Button(
            f_btns, 
            text="➕  Añadir", 
            command=self.controller.agregar_carpeta, 
            bg="#16a34a",
            fg="white",
            font=("Arial", 8, "bold"),
            relief=tk.FLAT,
            cursor="hand2",
            padx=8, pady=3
        ).pack(side=tk.LEFT)
        
        tk.Button(
            f_btns, 
            text="✕  Quitar", 
            command=self.controller.quitar_carpeta,
            bg="#dc2626",
            fg="white",
            font=("Arial", 8, "bold"),
            relief=tk.FLAT,
            cursor="hand2",
            padx=8, pady=3
        ).pack(side=tk.LEFT, padx=3)

        tk.Button(
            f_btns,
            text="🗑  Limpiar todo",
            command=self.controller.limpiar_carpetas,
            bg="#7f1d1d",
            fg="white",
            font=("Arial", 8, "bold"),
            relief=tk.FLAT,
            cursor="hand2",
            padx=8, pady=3
        ).pack(side=tk.LEFT)

    def _crear_seccion_excepciones(self, parent):
        """Crea la sección de excepciones (no procesar / no copiar)"""
        f_exc = tk.Frame(parent)
        f_exc.pack(fill=tk.X, pady=3)

        # No procesar
        f1 = tk.LabelFrame(
            f_exc, 
            text="🚫 No procesar (formato/pdf)", 
            font=("Arial", 8, "bold"),
            padx=4,
            pady=3
        )
        f1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 2))
        
        self.text_no_process = tk.Text(f1, height=2, font=("Arial", 8))
        self.text_no_process.pack(fill=tk.BOTH)

        # No copiar
        f2 = tk.LabelFrame(
            f_exc, 
            text="🚫 Ignorados (archivo/carpeta)", 
            font=("Arial", 8, "bold"),
            padx=4,
            pady=3
        )
        f2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(2, 0))
        
        self.text_no_copy = tk.Text(f2, height=2, font=("Arial", 8))
        self.text_no_copy.pack(fill=tk.BOTH)

    def _crear_seccion_destino(self, parent):
        """Crea la sección de carpeta destino"""
        frame = tk.Frame(parent)
        frame.pack(fill=tk.X, pady=3)
        
        tk.Label(
            frame, 
            text="💾 Destino:", 
            font=("Arial", 9, "bold")
        ).pack(side=tk.LEFT)
        
        self.entry_destino = tk.Entry(frame, font=("Arial", 8))
        self.entry_destino.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Cargar destino por defecto desde config si existe
        if hasattr(self.controller, 'config_manager'):
            destino_default = self.controller.config_manager.get_str('USER', 'last_destination')
            if destino_default:
                self.entry_destino.insert(0, destino_default)
        
        tk.Button(
            frame, 
            text="📂", 
            command=self.controller.seleccionar_destino,
            bg="#2563eb",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=6, pady=3
        ).pack(side=tk.LEFT)

        tk.Button(
            frame,
            text="🗑",
            command=self.controller.limpiar_destino,
            bg="#dc2626",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=6, pady=3
        ).pack(side=tk.LEFT, padx=(3, 0))

    def _crear_boton_empezar(self, parent):
        """Crea el botón principal de inicio"""
        self.btn_empezar = tk.Button(
            parent, 
            text="🚀  EMPEZAR CONVERSIÓN", 
            font=("Arial", 10, "bold"), 
            bg="#16a34a",
            fg="white", 
            command=self.controller.empezar_proceso, 
            relief=tk.FLAT,
            cursor="hand2",
            pady=6
        )
        self.btn_empezar.pack(fill=tk.X, pady=4)

    def _crear_seccion_progreso(self, parent):
        """Crea la sección de barra de progreso"""
        self.label_progreso = tk.Label(
            parent, 
            text="⏳ Esperando...", 
            font=("Arial", 8)
        )
        self.label_progreso.pack(anchor=tk.W)
        
        # Configurar estilo de barra de progreso
        style = ttk.Style()
        style.theme_use('default')
        style.configure(PROGRESS_BAR_STYLE, **PROGRESS_BAR_COLORS)
        
        self.progress_bar = ttk.Progressbar(
            parent, 
            mode='determinate', 
            style=PROGRESS_BAR_STYLE
        )
        self.progress_bar.pack(fill=tk.X)

    def _crear_seccion_log(self, parent):
        """Crea la sección de log de progreso"""
        self.log_text = scrolledtext.ScrolledText(
            parent, 
            height=6, 
            font=("Consolas", 8)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(3, 0))

    def _cargar_autor_desde_archivo(self):
        """Carga el contenido de autor.txt como valor por defecto"""
        try:
            ruta_autor = os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                "autor.txt"
            )
            if os.path.exists(ruta_autor):
                with open(ruta_autor, 'r', encoding='utf-8') as f:
                    return f.read().strip()
        except Exception as e:
            print(f"No se pudo cargar autor.txt: {e}")
        return ""

    # ====
    # MÉTODOS PÚBLICOS PARA ACTUALIZAR LA INTERFAZ
    # ====

    def log(self, mensaje):
        """Agrega un mensaje al log"""
        self.log_text.insert(tk.END, mensaje + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def limpiar_log(self):
        """Limpia el contenido del log"""
        self.log_text.delete(1.0, tk.END)

    def actualizar_progreso(self, valor, texto=None):
        """Actualiza la barra de progreso"""
        self.progress_bar['value'] = valor
        if texto:
            self.label_progreso.config(text=f"⏳ {texto}")
        self.root.update_idletasks()

    def actualizar_label_logo(self, texto, exitoso=True):
        """Método obsoleto - mantenido por compatibilidad"""
        pass  # El label ya no existe, se usa solo el canvas preview

    def mostrar_preview_logo(self, ruta_logo):
        """Muestra una previsualización del logo en el Canvas - CENTRADO"""
        if self.canvas_logo_preview is None:
            return
            
        try:
            img = Image.open(ruta_logo)
            
            # Forzar actualización del canvas para obtener dimensiones reales
            self.canvas_logo_preview.update()
            
            # Obtener dimensiones reales del canvas
            canvas_w = self.canvas_logo_preview.winfo_width()
            canvas_h = self.canvas_logo_preview.winfo_height()
            
            # Si no tiene dimensiones, usar las configuradas
            if canvas_w <= 1:
                canvas_w = self.canvas_preview_width
            if canvas_h <= 1:
                canvas_h = self.canvas_preview_height
            
            # Calcular dimensiones con padding mínimo
            max_w = canvas_w - 10
            max_h = canvas_h - 10
            
            # Calcular ratio manteniendo aspecto - maximizar el tamaño
            ratio = min(max_w / img.width, max_h / img.height)
            new_w = int(img.width * ratio)
            new_h = int(img.height * ratio)
            
            # Redimensionar imagen
            img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(img_resized)
            
            # Limpiar canvas
            self.canvas_logo_preview.delete("all")
            self.canvas_logo_preview.config(bg="white")
            
            # Calcular centro REAL del canvas
            center_x = canvas_w / 2.0
            center_y = canvas_h / 2.0
            
            # Crear imagen CENTRADA con coordenadas flotantes
            self.canvas_image_id = self.canvas_logo_preview.create_image(
                center_x, center_y,
                image=self.logo_photo,
                anchor=tk.CENTER
            )
            
        except Exception as e:
            if self.canvas_logo_preview:
                self.canvas_logo_preview.delete("all")
                self.canvas_logo_preview.config(bg=COLOR_LOGO_BG)
                center_x = self.canvas_preview_width / 2.0
                center_y = self.canvas_preview_height / 2.0
                self.canvas_logo_preview.create_text(
                    center_x, center_y,
                    text=f"❌ Error al cargar\n{str(e)[:30]}",
                    fill="#ff0000",
                    font=("Arial", 9)
                )

    def limpiar_preview_logo(self):
        """Limpia la previsualización del logo"""
        if self.canvas_logo_preview is None:
            return
            
        self.logo_photo = None
        self.canvas_image_id = None
        self.canvas_logo_preview.delete("all")
        self.canvas_logo_preview.config(bg=COLOR_LOGO_BG)
        self.canvas_logo_preview.create_text(
            self.canvas_preview_width // 2,
            self.canvas_preview_height // 2,
            text="Arrastra aquí el logo\n(PNG/JPG)",
            fill="#9999",
            font=("Arial", 10),
            tags="placeholder"
        )

    def deshabilitar_boton_empezar(self):
        """Deshabilita el botón de empezar"""
        self.btn_empezar.config(state=tk.DISABLED, bg=COLOR_DISABLED)

    def habilitar_boton_empezar(self):
        """Habilita el botón de empezar"""
        self.btn_empezar.config(state=tk.NORMAL, bg="#16a34a")

    # ====
    # MÉTODOS PARA OBTENER DATOS DE LA INTERFAZ
    # ====

    def obtener_palabras_prohibidas(self):
        """Obtiene la lista de palabras prohibidas (excepciones de procesamiento)"""
        texto = self.text_no_process.get("1.0", tk.END).strip()
        if not texto:
            return []
        palabras = [linea.strip().lower() for linea in texto.split('\n') if linea.strip()]
        return palabras

    def obtener_autor(self):
        """Obtiene el nombre del autor"""
        return self.entry_autor.get().strip()

    def obtener_carpeta_destino(self):
        """Obtiene la ruta de la carpeta destino"""
        return os.path.normpath(self.entry_destino.get().strip())

    def obtener_opciones(self):
        """Obtiene las opciones básicas (para compatibilidad con código anterior)"""
        return {
            'encabezado': self.var_add_logo.get() or self.var_add_folder_code.get() or self.var_add_header_line.get(),
            'pie_pagina': self.var_add_footer_line.get() or self.var_add_author.get() or self.var_add_page_number.get(),
            'copiar_archivos': self.var_copy_attachments.get(),
            'respetar_estructura': self.var_respect_structure.get()
        }

    def obtener_opciones_completas(self):
        """Captura el estado actual de toda la interfaz para enviarlo al controlador"""
        return {
            # Encabezado y Pie
            'add_logo': self.var_add_logo.get(),
            'add_folder_code': self.var_add_folder_code.get(),
            'add_header_line': self.var_add_header_line.get(),
            'add_footer_line': self.var_add_footer_line.get(),
            'add_author': self.var_add_author.get(),
            'add_page_number': self.var_add_page_number.get(),
            'autor_nombre': self.entry_autor.get().strip(),

            # Opciones de Copia
            'respect_structure': self.var_respect_structure.get(),
            'copy_attachments': self.var_copy_attachments.get(),
            'save_modified_dest': self.var_save_modified_dest.get(),
            'copy_as_pdf': self.var_copy_as_pdf.get(),

            # Extensiones
            'process_docx': self.var_process_docx.get(),
            'process_docm': self.var_process_docm.get(),

            # Rutas y Listas
            'carpetas': list(self.listbox_carpetas.get(0, tk.END)),
            'destino': self.entry_destino.get().strip(),
            'excepciones_procesar': self.text_no_process.get("1.0", tk.END).strip(),
            'excepciones_copiar': self.text_no_copy.get("1.0", tk.END).strip()
        }

    # ====
    # MÉTODOS PARA GESTIONAR CARPETAS
    # ====

    def establecer_carpeta_destino(self, ruta):
        """Establece la ruta de la carpeta destino"""
        self.entry_destino.delete(0, tk.END)
        self.entry_destino.insert(0, ruta)

    def agregar_carpeta_a_lista(self, carpeta):
        """Agrega una carpeta a la lista"""
        self.listbox_carpetas.insert(tk.END, carpeta)

    def quitar_carpeta_de_lista(self, index):
        """Quita una carpeta de la lista por índice"""
        self.listbox_carpetas.delete(index)

    def obtener_seleccion_carpeta(self):
        """Obtiene el índice de la carpeta seleccionada"""
        seleccion = self.listbox_carpetas.curselection()
        return seleccion[0] if seleccion else None

    def limpiar_lista_carpetas(self):
        """Limpia toda la lista de carpetas"""
        self.listbox_carpetas.delete(0, tk.END)

    # ====
    # DIÁLOGOS
    # ====

    def mostrar_error(self, titulo, mensaje):
        """Muestra un diálogo de error"""
        from tkinter import messagebox
        messagebox.showerror(titulo, mensaje)

    def mostrar_info(self, titulo, mensaje):
        """Muestra un diálogo de información"""
        from tkinter import messagebox
        messagebox.showinfo(titulo, mensaje)

    def mostrar_pregunta(self, titulo, mensaje):
        """Muestra un diálogo de pregunta (Sí/No)"""
        from tkinter import messagebox
        return messagebox.askyesno(titulo, mensaje)

    def ejecutar(self):
        """Inicia el loop principal de la interfaz"""
        self.root.mainloop()
                
                
                
                
                
                
                
                
    def _crear_seccion_opciones_detalladas(self, parent):
        """Crea las secciones de opciones detalladas con checkboxes"""
        # Encabezado y Pie de Página
        f_hp = tk.LabelFrame(
            parent, 
            text="📄 Encabezado y Pie de Página", 
            font=("Arial", 8, "bold"), 
            padx=6, 
            pady=3
        )
        f_hp.pack(fill=tk.X, pady=3)

        opts_hp = [
            ("Logo en encabezado", self.var_add_logo),
            ("Código carpeta", self.var_add_folder_code),
            ("Línea encabezado", self.var_add_header_line),
            ("Línea pie página", self.var_add_footer_line),
            ("Añadir Autor", self.var_add_author),
            ("Número página", self.var_add_page_number)
        ]
        for i, (txt, var) in enumerate(opts_hp):
            tk.Checkbutton(
                f_hp, 
                text=txt, 
                variable=var, 
                font=("Arial", 8)
            ).grid(row=i//2, column=i%2, sticky=tk.W, padx=3, pady=1)

        # Opciones de Copia
        f_cp = tk.LabelFrame(
            parent, 
            text="📂 Opciones de Copia", 
            font=("Arial", 8, "bold"), 
            padx=6, 
            pady=3
        )
        f_cp.pack(fill=tk.X, pady=3)

        opts_cp = [
            ("Respetar estructura", self.var_respect_structure),
            ("Copiar anexos", self.var_copy_attachments),
            ("Guardar modificado en destino", self.var_save_modified_dest),
            ("Copiar como PDF", self.var_copy_as_pdf),
            ("Renombrar con código carpeta", self.var_auto_rename)
        ]
        for i, (txt, var) in enumerate(opts_cp):
            tk.Checkbutton(
                f_cp, 
                text=txt, 
                variable=var, 
                font=("Arial", 8)
            ).grid(row=i//2, column=i%2, sticky=tk.W, padx=3, pady=1)


        
        
        
        
    def test_dialogo_simple(self):
        """Método de prueba para verificar que los diálogos funcionan"""
        from tkinter import simpledialog
        
        print("TEST: Mostrando diálogo de prueba...")
        resultado = simpledialog.askstring(
            "Prueba",
            "Escribe algo de prueba:",
            parent=self.root
        )
        print(f"TEST: Resultado = {repr(resultado)}")
        self.log(f"Resultado del test: {resultado}")
#libreias que utiliza el programa, son necesarias tener instaladas para la ejecucion por codigo
import os
import shutil
import subprocess
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from collections import defaultdict
from threading import Thread, Event
from PIL import Image, ImageTk
from datetime import datetime
import tempfile, subprocess
import io
import tempfile
import os
from pathlib import Path
import fitz  
from PIL import Image, ImageTk
import pygame
import tkinter as tk
from threading import Thread
import sys
import ctypes
from ctypes import wintypes
from PIL import Image, ImageTk
import io
from matplotlib.colors import to_hex
pygame.init()
pygame.mixer.init()

# Opcionales: instala si quieres preview completo
try:
    import pygame
    pygame.mixer.init()
except ImportError:
    pygame = None

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

try:
    import docx
except ImportError:
    docx = None

try:
    import cv2
except ImportError:
    cv2 = None
try:
    import pygame
    pygame.mixer.init()
except ImportError:
    pygame = None
try:
    import fitz 
except ImportError:
    fitz = None
try:
    import docx
except ImportError:
    docx = None
try:
    from pptx import Presentation
except ImportError:
    Presentation = None
try:
    import openpyxl
except ImportError:
    openpyxl = None

#paleta de colorores de la ui
PALETTE = ["#FFE5B4", "#FAD896", "#F5CB78", "#F0BE5A", "#EBC13C", "#E6AE1E", "#DD9D00", "#D38C00", "#C97B00", "#BF6A00"]

#clase principal
class FileManagerApp:
    #funcion de inicio del programa
    def __init__(self, root):
        self.root = root
        self.root.title("File Flow Tracker")
        self.root.configure(bg=PALETTE[0])
        self.cancel_event = Event()
        self.dir_path = None
        self.files = []
        self.file_types = defaultdict(list)
        self.current_sort = {}
        self.apply_styles()
        self.setup_ui()
        self.root.state('zoomed') 
        pygame.init()
        pygame.mixer.init()
        self.audio_channel = pygame.mixer.Channel(0)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    #Funcion al cerrar la ui
    def on_close(self):
        # 1) Señala a los hilos que deben parar
        self.cancel_event.set()

        # 2) Si estás reproduciendo audio:
        try:
            self.audio_channel.stop()
            pygame.mixer.quit()
        except:
            pass

        # 3) Destruye la ventana
        self.root.destroy()

        # 4) Y asegúrate de salir del intérprete
        sys.exit(0)

    #aplicar los stilos de la paleta de colores a la ui
    def apply_styles(self):
        style = ttk.Style(self.root)
        style.theme_use('clam')
        style.configure('TFrame', background=PALETTE[0])
        style.configure('TLabel', background=PALETTE[0], foreground=PALETTE[-1])
        style.configure('TButton', background=PALETTE[4], foreground=PALETTE[-1])
        style.map('TButton', background=[('active', PALETTE[5])])
        style.configure('Vertical.TScrollbar', background=PALETTE[2], troughcolor=PALETTE[0])
        style.configure('Horizontal.TScrollbar', background=PALETTE[2], troughcolor=PALETTE[0])
        style.configure('Treeview', background=PALETTE[1], fieldbackground=PALETTE[1], foreground=PALETTE[-2])
        style.configure('Treeview.Heading', background=PALETTE[3], foreground=PALETTE[-1])
        style.configure("TCombobox", fieldbackground=PALETTE[1], background=PALETTE[1], foreground=PALETTE[-1])

        # Estilo para el Labelframe en general
        style.configure('TLabelframe', 
            background=PALETTE[0],    # fondo del contenedor
            bordercolor=PALETTE[3],
            anchor='center'    # color del borde (opcional)
        )
        # Estilo para el texto del título
        style.configure('TLabelframe.Label',
            background=PALETTE[0],    # mismo fondo que el frame
            foreground=PALETTE[-1],   # color del texto
            anchor='center'           # centrar el texto
        )

    def setup_ui(self):
        main = ttk.Frame(self.root)
        main.grid(row=0, column=0, sticky='nsew')
        self.root.rowconfigure(0, weight=1); self.root.columnconfigure(0, weight=1)

        # panel de control
        ctrl = ttk.Frame(main)
        ctrl.grid(row=0, column=0, sticky='ew', padx=10, pady=5)
        ttk.Button(ctrl, text="Seleccionar Carpeta", command=self.on_select).grid(row=0, column=0)
        ttk.Button(ctrl, text="Exportar", command=self.export_data).grid(row=0, column=1, padx=5)
        ttk.Label(ctrl, text="Filtro por tipo:").grid(row=0, column=2, padx=5)
        self.filter_cb = ttk.Combobox(ctrl)
        self.filter_cb.grid(row=0, column=3)
        self.filter_cb.bind("<<ComboboxSelected>>", lambda e: self.populate_extra_tree())
        self.filter_cb.bind("<KeyRelease>", self.on_filter_key)

        # Panel para las tablas
        pane = ttk.PanedWindow(main, orient=tk.HORIZONTAL)
        pane.grid(row=1, column=0, sticky='nsew', padx=10, pady=5)
        pane.rowconfigure(0, weight=1)
        pane.columnconfigure(0, weight=1)
        pane.columnconfigure(1, weight=3)
        pane.columnconfigure(2, weight=2)
        main.rowconfigure(1, weight=3)


        # Summary tree
        summary_frame = ttk.Labelframe(
            pane,
            text="Resumen por Tipo de los Archivos",
            style='TLabelframe',
            labelanchor='n'  
        )
        #tabla de tipo archivo mas pesado
        summary_frame.grid_rowconfigure(0, weight=1); summary_frame.grid_columnconfigure(0, weight=1)
        cols_sum = ("Tipo","Tamaño_MB")
        self.summary_tree = ttk.Treeview(summary_frame, columns=cols_sum, show='headings')
        for c, t in zip(cols_sum, ["Tipo","Tamaño (MB)"]):
            self.summary_tree.heading(c, text=t,
                command=lambda c=c: self.sort_tree(self.summary_tree,c,c=="Tamaño_MB"))
            self.summary_tree.column(c, anchor='center')
        vsb=ttk.Scrollbar(summary_frame, orient=tk.VERTICAL, command=self.summary_tree.yview)
        hsb=ttk.Scrollbar(summary_frame, orient=tk.HORIZONTAL, command=self.summary_tree.xview)
        self.summary_tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.summary_tree.grid(row=0,column=0,sticky='nsew')
        vsb.grid(row=0,column=1,sticky='ns'); hsb.grid(row=1,column=0,sticky='ew')
        self.summary_tree.bind("<Double-1>", self.open_in_explorer_folder)
        pane.add(summary_frame, weight=1)

        # titulo del panel de las tablas
        detail_frame = ttk.Labelframe(
        pane,
        text="Listado General de los Archivos de la Carpeta/Disco",
        style='TLabelframe',
        labelanchor='n'
        )
        # tabla de resumen general de los archivos
        detail_frame = ttk.Labelframe(pane, text="Resumen General de los Archivos de la Carpeta/Disco")
        detail_frame.grid_rowconfigure(0, weight=1); detail_frame.grid_columnconfigure(0, weight=1)
        cols_det = ("Archivo","Tamaño_MB","Tipo","Accion")
        self.extra_tree = ttk.Treeview(detail_frame, columns=cols_det, show='headings')
        for c,txt,num in [("Archivo","Archivo",False),("Tamaño_MB","Tamaño (MB)",True),
                          ("Tipo","Tipo",False),("Accion","Accion",False)]:
            self.extra_tree.heading(c,text=txt,
                command=lambda c=c,n=num: self.sort_tree(self.extra_tree,c,n))
            self.extra_tree.column(c,anchor='center')
        vsb2=ttk.Scrollbar(detail_frame, orient=tk.VERTICAL, command=self.extra_tree.yview)
        hsb2=ttk.Scrollbar(detail_frame, orient=tk.HORIZONTAL, command=self.extra_tree.xview)
        self.extra_tree.configure(yscroll=vsb2.set, xscroll=hsb2.set)
        self.extra_tree.grid(row=0,column=0,sticky='nsew')
        vsb2.grid(row=0,column=1,sticky='ns'); hsb2.grid(row=1,column=0,sticky='ew')
        self.extra_tree.bind("<Double-1>", self.open_in_explorer_file)
        self.extra_tree.bind("<ButtonRelease-1>", self.on_action_click)
        pane.add(detail_frame, weight=3)
        
        # Tabla de subcarpetas mas pesadas
        folder_frame = ttk.Labelframe(pane, text="Subcarpetas Pesadas", style='TLabelframe', labelanchor='n')
        folder_frame.grid_rowconfigure(0, weight=1)
        folder_frame.grid_columnconfigure(0, weight=1)
        pane.add(folder_frame, weight=2) 
        cols_folders = ("Carpeta","Tamaño_MB","Ubicación")
        self.folder_tree = ttk.Treeview(folder_frame, columns=cols_folders, show='headings')
        for c, t in zip(cols_folders, ["Carpeta","Tamaño (MB)","Ubicación"]):
            self.folder_tree.heading(c, text=t,
                command=lambda c=c: self.sort_tree(self.folder_tree, c, c=="Tamaño_MB"))
            self.folder_tree.column(c, anchor='center')
        # Scrollbars para las tablas:
        vsb_f = ttk.Scrollbar(folder_frame, orient=tk.VERTICAL, command=self.folder_tree.yview)
        hsb_f = ttk.Scrollbar(folder_frame, orient=tk.HORIZONTAL, command=self.folder_tree.xview)
        self.folder_tree.configure(yscroll=vsb_f.set, xscroll=hsb_f.set)
        self.folder_tree.grid(row=0, column=0, sticky='nsew')
        vsb_f.grid(row=0, column=1, sticky='ns')
        hsb_f.grid(row=1, column=0, sticky='ew')
        self.folder_tree.bind("<Double-1>", lambda e: subprocess.run(['explorer', '/select,', self.folder_tree.selection()[0]]))
        

        # ————— Chart + Preview pane —————
        cp_pane = ttk.PanedWindow(main, orient=tk.HORIZONTAL)
        cp_pane.grid(row=2, column=0, sticky='nsew', padx=10, pady=5)
        main.rowconfigure(2, weight=2)
        main.columnconfigure(0, weight=1)

        # Contenedor del gráfico
        self.chart_frame = ttk.Labelframe(
            cp_pane,
            text="Archivos más pesados",
            style='TLabelframe',
            labelanchor='n'
        )
        # peso del grafico
        self.chart_frame.grid_rowconfigure(0, weight=1)
        self.chart_frame.grid_columnconfigure(0, weight=1)
        cp_pane.add(self.chart_frame, weight=3)

        #Contenedor de la previsualización
        self.preview_frame = ttk.Labelframe(
            cp_pane,
            text="Previsualización",
            style='TLabelframe',
            labelanchor='n'
        )
        self.preview_frame.grid_rowconfigure(0, weight=1)
        self.preview_frame.grid_columnconfigure(0, weight=1)
        cp_pane.add(self.preview_frame, weight=2)

        # Ahora creamos la tabla de subcarpetas pesadas
        cols_folders = ("Carpeta","Tamaño_MB","Ubicación")
        self.folder_tree = ttk.Treeview(folder_frame, columns=cols_folders, show='headings')
        for c, t in zip(cols_folders, ["Carpeta","Tamaño (MB)","Ubicación"]):
            self.folder_tree.heading(c, text=t,
                command=lambda col=c: self.sort_tree(self.folder_tree, col, col=="Tamaño_MB"))
            self.folder_tree.column(c, anchor='center')
        # Scrollbars
        vsb_f = ttk.Scrollbar(folder_frame, orient=tk.VERTICAL, command=self.folder_tree.yview)
        hsb_f = ttk.Scrollbar(folder_frame, orient=tk.HORIZONTAL, command=self.folder_tree.xview)
        self.folder_tree.configure(yscroll=vsb_f.set, xscroll=hsb_f.set)
        # Layout
        self.folder_tree.grid(row=0, column=0, sticky='nsew')
        vsb_f.grid(row=0, column=1, sticky='ns')
        hsb_f.grid(row=1, column=0, sticky='ew')
        # Doble clic para abrir el Explorador
        self.folder_tree.bind(
            "<Double-1>",
            lambda e: subprocess.run(['explorer', '/select,', self.folder_tree.selection()[0]])
        )

        # ——— Chart + Preview pane ———
        cp_pane = ttk.PanedWindow(main, orient=tk.HORIZONTAL)
        cp_pane.grid(row=2, column=0, sticky='nsew', padx=10, pady=5)
        main.rowconfigure(2, weight=2)
        main.columnconfigure(0, weight=1)

        # Chart frame (existente)
        self.chart_frame = ttk.Frame(cp_pane)
        cp_pane.add(self.chart_frame, weight=2)

        # Preview frame (nuevo)
        self.preview_frame = ttk.Labelframe(cp_pane, text="Previsualización", style='TLabelframe', labelanchor='n')
        self.preview_frame.grid_rowconfigure(0, weight=1)  # fila de contenido
        self.preview_frame.grid_rowconfigure(1, weight=0)  # fila de footer
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.configure(width=500, height=500)
        cp_pane.add(self.preview_frame, weight=1)
        for w in self.preview_frame.winfo_children():
            w.destroy()

        # Bind selección de extra_tree
        self.extra_tree.bind("<<TreeviewSelect>>", self.on_file_select)
         
         #filtro
    def on_filter_key(self, event):
        txt = self.filter_cb.get().lower()
        vals = ['Todos'] + [e for e in self.file_types if txt in e]
        self.filter_cb['values'] = vals
        #selecionar archivo
    def on_file_select(self, event):
        sel = self.extra_tree.selection()
        if not sel: return
        path = Path(sel[0])
        ext = path.suffix.lower()
        self.preview_file(path, ext)
    #funcion para archivos de audio
    def play_audio(self, path: Path):
        import pygame
        # Primero, intenta pyg­ame para formatos soportados
        try:
            sound = pygame.mixer.Sound(str(path))
            self.audio_channel.stop()
            self.audio_channel.play(sound)
            return
        except Exception:
            pass
        # Fallback: abre con la app predeterminada de Windows
        try:
            os.startfile(str(path))
        except Exception as e:
            messagebox.showerror("Error audio", f"No se pudo reproducir el audio:\n{e}")
    #funcion para detener la reproduccion del archivo de audio
    def stop_audio(self):
        try:
            self.audio_channel.stop()
        except:
            pass
    #funcion para obtener el icono de los archivos
    def get_file_icon(self, path: Path, size=64):
        """
        Devuelve un PhotoImage con el icono de Windows asociado a `path`.
        Requiere pywin32 (win32gui, win32ui, win32con).
        """
        import ctypes
        from ctypes import wintypes
        import win32gui, win32ui, win32con

        # Flags para SHGetFileInfo
        SHGFI_ICON              = 0x100
        SHGFI_USEFILEATTRIBUTES = 0x10
        SHGFI_LARGEICON         = 0x0
        SHGFI_SMALLICON         = 0x1

        class SHFILEINFOW(ctypes.Structure):
            _fields_ = [
                ("hIcon",       wintypes.HICON),
                ("iIcon",       wintypes.INT),
                ("dwAttributes",wintypes.DWORD),
                ("szDisplayName", wintypes.WCHAR * 260),
                ("szTypeName",    wintypes.WCHAR * 80),
            ]

        flags = SHGFI_ICON | SHGFI_USEFILEATTRIBUTES
        flags |= SHGFI_LARGEICON if size > 32 else SHGFI_SMALLICON

        shfi = SHFILEINFOW()
        res = ctypes.windll.shell32.SHGetFileInfoW(
            str(path),
            0,
            ctypes.byref(shfi),
            ctypes.sizeof(shfi),
            flags
        )
        if res == 0 or not shfi.hIcon:
            return None

        # Preparamos DCs
        hdc_screen = win32gui.GetDC(0)
        hdc_mem    = win32gui.CreateCompatibleDC(hdc_screen)
        hbm        = win32gui.CreateCompatibleBitmap(hdc_screen, size, size)
        win32gui.SelectObject(hdc_mem, hbm)
        # Dibujamos el icono en el bitmap
        win32gui.DrawIconEx(hdc_mem, 0, 0, shfi.hIcon, size, size, 0, 0, win32con.DI_NORMAL)

        # Convertimos el HBITMAP a objeto Bitmap de win32ui
        bmp = win32ui.CreateBitmapFromHandle(hbm)
        # Ahora obtenemos los bytes
        bmpstr = bmp.GetBitmapBits(True)
        bmpinfo = bmp.GetInfo()

        # Creamos la imagen PIL
        img = Image.frombuffer(
            'RGBA',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRA', 0, 1
        )

        # Limpieza
        win32gui.DeleteObject(hbm)
        win32gui.DeleteDC(hdc_mem)
        win32gui.ReleaseDC(0, hdc_screen)
        ctypes.windll.user32.DestroyIcon(shfi.hIcon)

        # Redimensionamos y empaquetamos en PhotoImage
        img = img.resize((size, size), Image.LANCZOS)
        return ImageTk.PhotoImage(img)

        #para obtener la primera pagina de archivos office
    def office_to_image(self, path: Path, container) -> bool:
        
        """
        Convierte primera página de docx/pptx/xlsx y la pinta dentro de `container`.
        Devuelve True si dibujó algo, False en caso contrario.
        """
        import win32com.client, tempfile, os, io
        try:
            ext = path.suffix.lower()
            tmp_pdf = Path(tempfile.gettempdir()) / (path.stem + "_preview.pdf")
            #para archivos de word
            if ext == '.docx':
                word = win32com.client.Dispatch('Word.Application')
                doc = word.Documents.Open(str(path))
                doc.SaveAs(str(tmp_pdf), FileFormat=17)  # 17 = wdFormatPDF
                doc.Close(); word.Quit()
            #para archivos power point
            elif ext == '.pptx':
                ppt = win32com.client.Dispatch('PowerPoint.Application')
                pres = ppt.Presentations.Open(str(path), WithWindow=False)
                pres.SaveAs(str(tmp_pdf), FileFormat=32)  # 32 = ppSaveAsPDF
                pres.Close(); ppt.Quit()
            #para archivos excel
            elif ext == '.xlsx':
                excel = win32com.client.Dispatch('Excel.Application')
                wb = excel.Workbooks.Open(str(path))
                wb.ExportAsFixedFormat(0, str(tmp_pdf))   # 0 = PDF
                wb.Close(False); excel.Quit()

            else:
                return False

            # ahora renderizamos la primera página:
            doc = fitz.open(str(tmp_pdf))
            page = doc.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(0.5,0.5))
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            img.thumbnail((480,480))
            photo = ImageTk.PhotoImage(img)

            lbl = ttk.Label(container, image=photo)
            lbl.image = photo
            lbl.pack(anchor='center', expand=True, pady=10)

            try: tmp_pdf.unlink()
            except: pass
            return True
        except Exception as e:
            print("office_to_image error:", e)
            return False
    #funcion para la previzualiacion de los archivos
    def preview_file(self, path: Path, ext: str):
        # 1) Limpia preview_frame
        for w in self.preview_frame.winfo_children():
            w.destroy()

        # 2) Crea un frame “container” que centre todo con pack
        container = ttk.Frame(self.preview_frame)
        container.pack(expand=True, fill='both')
        container.columnconfigure(0, weight=1)

        # 3) Metadatos
        size = path.stat().st_size / (1024**2)
        mtime = path.stat().st_mtime
        date = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
        footer = (
            f"Nombre: {path.name}\n"
            f"Tamaño: {size:.2f} MB\n"
            f"Modificado: {date}\n"
            f"Tipo: {ext}"
        )

        # 4) Contenido principal (imagen, vídeo, audio, etc.)
        #para imagenes
        if ext in ('.png','.jpg','.jpeg','.gif','.bmp','.webp','.tiff','.ico'):
            try:
                img = Image.open(path)
                img.thumbnail((480,480))
                photo = ImageTk.PhotoImage(img)
                lbl = ttk.Label(container, image=photo)
                lbl.image = photo
                lbl.pack(anchor='center', expand=True, pady=10)
            except:
                pass
        #para videos
        elif ext in ('.mp4','.mkv','.avi','.mov') and cv2:
            try:
                cap = cv2.VideoCapture(str(path))
                ret, frame = cap.read()
                cap.release()
                if ret:
                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    img = Image.fromarray(frame)
                    img.thumbnail((480,480))
                    photo = ImageTk.PhotoImage(img)
                    lbl = ttk.Label(container, image=photo)
                    lbl.image = photo
                    lbl.pack(anchor='center', expand=True, pady=10)
            except:
                pass
        #para audios
        elif ext in ('.mp3','.wav','.m4a','.flac'):
            ttk.Label(container, text=path.name).pack(anchor='center', pady=(10,5))
            btnf = ttk.Frame(container)
            btnf.pack(anchor='center', pady=(0,10))
            ttk.Button(btnf, text="▶️", command=lambda p=path: self.play_audio(p)).pack(side=tk.LEFT, padx=5)
            ttk.Button(btnf, text="⏹️", command=self.stop_audio).pack(side=tk.LEFT, padx=5)
        #para pdfs
        elif ext == '.pdf':
            if fitz:
                try:
                    doc = fitz.open(str(path))
                    page = doc.load_page(0)
                    pix = page.get_pixmap(matrix=fitz.Matrix(0.5,0.5))
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    photo = ImageTk.PhotoImage(img)
                    lbl = ttk.Label(container, image=photo)
                    lbl.image = photo
                    lbl.pack(anchor='center', expand=True, pady=10)
                except:
                    pass
            else:
                text = "[No se puede mostrar PDF]"
                if PyPDF2:
                    try:
                        reader = PyPDF2.PdfReader(str(path))
                        text = reader.pages[0].extract_text()[:500] or text
                    except:
                        pass
                txt = tk.Text(container, wrap='word')
                txt.insert('1.0', text)
                txt.configure(state='disabled')
                txt.pack(expand=True, fill='both', padx=10, pady=10)
        #para archivos de office
        elif ext in ('.docx','.pptx','.xlsx'):
            if self.office_to_image(path, container):
                # office_to_image ya pintó dentro de container
                pass
            else:
                lbl_icon = ttk.Label(container, text="📄", font=("Arial", 64))
                lbl_icon.pack(anchor='center', pady=20)
        #para archivos de texto
        elif ext in ('.txt','.html','.json','.xml','.css','.js'):
            snippet = "[Error leyendo texto]"
            try:
                with open(path, encoding='utf8', errors='ignore') as f:
                    snippet = "".join(f.readlines()[:20])
            except:
                pass
            txt = tk.Text(container, wrap='word')
            txt.insert('1.0', snippet)
            txt.configure(state='disabled')
            txt.pack(expand=True, fill='both', padx=10, pady=10)

        # Resto de tipos…
        else:
            # extraemos icono real o emoji
            icon = self.get_file_icon(path, size=64)
            if icon:
                lbl_icon = ttk.Label(container, image=icon)
                lbl_icon.image = icon
            else:
                lbl_icon = ttk.Label(container, text="📄", font=("Arial", 64))
            lbl_icon.pack(anchor='center', pady=(20, 5))

            # Descripción única
            size_mb = path.stat().st_size / (1024**2)
            mtime    = path.stat().st_mtime
            date_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            desc = (
                f"Nombre: {path.name}\n"
                f"Tamaño: {size_mb:.2f} MB\n"
                f"Modificado: {date_str}\n"
                f"Tipo: {ext}"
            )
            lbl_desc = ttk.Label(container, text=desc, justify='center', wraplength=280)
            lbl_desc.pack(anchor='center', pady=(0, 10))

            # **Aquí cortamos** para que no llegue al footer global
            return

        # 5) Footer centrado
        ttk.Label(container, text=footer, justify='center').pack(anchor='center', pady=(0,15))

    #funcion para la seleccion de directorio
    def on_select(self):
        path = filedialog.askdirectory()
        if not path:
            return
        self.dir_path = path

        # ————— LIMPIAR UI INMEDIATAMENTE —————
        self.summary_tree.delete(*self.summary_tree.get_children())
        self.extra_tree.delete(*self.extra_tree.get_children())
        for w in self.chart_frame.winfo_children():
            w.destroy()
        self.filter_cb.set('')  # resetear combobox

        self.cancel_event.clear()
        self.show_progress_popup()
        Thread(target=self.scan_directory, daemon=True).start()
    #funcion para mostrar el progreso del analisis
    def show_progress_popup(self):
        self.popup = tk.Toplevel(self.root)
        self.popup.title("Escaneando carpeta...")
        self.popup.configure(bg=PALETTE[0])
        self.popup.geometry("300x100")
        # Esto hace que el popup capture eventos, pero no bloquea el loop:
        self.popup.transient(self.root)
        self.popup.grab_set()

        ttk.Label(self.popup, text="Escaneando...\nPuede tardar un momento").pack(pady=5)
        self.popup_pb = ttk.Progressbar(self.popup, orient=tk.HORIZONTAL, length=250, mode='determinate')
        self.popup_pb.pack(pady=5)
        ttk.Button(self.popup, text="Cancelar", command=self.cancel_scan).pack()

    #cancelar scaneo
    def cancel_scan(self):
        self.cancel_event.set()
        if hasattr(self, 'popup') and self.popup.winfo_exists():
            self.popup.grab_release()
            self.popup.destroy()
    #crear trewview
    def create_treeview(self, parent, columns, dbl_click):
        frame = ttk.Frame(parent)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=columns, show='headings')
        for c in columns:
            tree.heading(c, text=c, command=lambda col=c: self.sort_tree(tree, col, col == "Tamaño_MB"))
            tree.column(c, anchor='center')
        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree.bind("<Double-1>", dbl_click)
        parent.add(frame, weight=1)
        return tree

    #ventana emergente del proceso de analisis de la carpeta
    def start_load(self):
        self.cancel_event.clear()
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("Analizando...")
        self.progress_window.geometry("300x80")
        self.progress_bar = ttk.Progressbar(self.progress_window, mode='determinate')
        self.progress_bar.pack(padx=10, pady=10, fill='x')
        ttk.Button(self.progress_window, text="Cancelar", command=self.cancel_event.set).pack(pady=5)
        Thread(target=self.load_directory, daemon=True).start()

    #cargar directorio
    def load_directory(self):
        path = filedialog.askdirectory()
        if not path:
            self.progress_window.destroy()
            return
        self.dir_path = path
        self.scan_directory()

    #escanear directorio
    def scan_directory(self):
        # Vaciar datos internos antes de comenzar
        self.files.clear()
        self.file_types.clear()

        # Contar el total de ficheros
        total = 0
        for root, dirs, files in os.walk(self.dir_path):
            total += len(files)
            if self.cancel_event.is_set():
                return

        # Fijar el máximo de la barra de progreso en el hilo principal
        self.root.after(0, lambda: self.popup_pb.config(maximum=total))

        count = 0
        for root, dirs, files in os.walk(self.dir_path):
            if self.cancel_event.is_set():
                return
            for name in files:
                if self.cancel_event.is_set():
                    return
                f = Path(root) / name
                try:
                    size = f.stat().st_size / (1024**2)
                except Exception:
                    continue
                ext = f.suffix.lower() or 'Sin extensión'

                # Guardar en las listas internas
                self.files.append((f, size, ext))
                self.file_types[ext].append((f, size))

                # Actualizar progreso
                count += 1
                self.root.after(0, self.safe_update_progress, count)

        # Al acabar, cerrar popup y refrescar UI
        self.root.after(0, self.finish_scan)
    
    #finalizar scaneo dentro de la ventana emergente
    def finish_scan(self):
        if hasattr(self, 'popup') and self.popup.winfo_exists():
            self.popup.grab_release()
            self.popup.destroy()
        #  Siempre refrescamos la UI con el estado actual (incluso si cancelaste)
        self.update_ui()

        # Preparamos el flag para un posible próximo escaneo
        self.cancel_event.clear()

    #actualizar ui
    def update_ui(self):
        # 1) Limpiar los contenidos visuales
        # - Árboles
        self.summary_tree.delete(*self.summary_tree.get_children())
        self.extra_tree.delete(*self.extra_tree.get_children())
        self.folder_tree.delete(*self.folder_tree.get_children())
        # - Gráfico
        for w in self.chart_frame.winfo_children():
            w.destroy()
        # - Previsualización
        for w in self.preview_frame.winfo_children():
            w.destroy()

        # 2) Volver a poblar TODO
        self.show_pie_chart()
        self.populate_filter()
        self.populate_summary_tree()
        self.populate_extra_tree()

    def safe_update_progress(self, value):
        if hasattr(self, 'popup_pb') and self.popup_pb.winfo_exists():
            self.popup_pb.config(value=value)

    #filtro de tipo de archivo
    def populate_filter(self):
        sorted_ext = sorted(self.file_types.keys(), key=lambda e: sum(sz for _, sz in self.file_types[e]), reverse=True)
        self.filter_cb['values'] = ['Todos'] + sorted_ext
        self.filter_cb.set('Todos')

    #Actualiza la tabla de resumen segun el tipo seleccionado
    def populate_summary_tree(self):
        self.summary_tree.delete(*self.summary_tree.get_children())
        for ext, items in sorted(self.file_types.items(), key=lambda x: sum(sz for _, sz in x[1]), reverse=True):
            total = sum(sz for _, sz in items)
            self.summary_tree.insert('', 'end', iid=ext, values=(ext, f"{total:.5f}"))
            # Limpia (añade self.folder_tree al UI)
        self.folder_tree.delete(*self.folder_tree.get_children())
        for sub in Path(self.dir_path).iterdir():
            if sub.is_dir():
                # suma recursiva rápida
                total = sum(f.stat().st_size for f in sub.rglob('*') if f.is_file())/(1024**2)
                self.folder_tree.insert('', 'end', iid=str(sub), values=(sub.name, f"{total:.2f}", str(sub)))

    #acciones para eliminar o mover archivo
    def populate_extra_tree(self):
        self.extra_tree.delete(*self.extra_tree.get_children())
        sel = self.filter_cb.get()
        for f, size, ext in sorted(self.files, key=lambda x: x[1], reverse=True):
            if sel == 'Todos' or ext == sel:
                action_text = "Eliminar | Mover"
                # usa la ruta como iid
                self.extra_tree.insert(
                    '', 'end',
                    iid=str(f),
                    values=(f.name, f"{size:.5f}", ext, "Eliminar | Mover")
                )
    #funcion que se ejecuta al clickear un archivo en la tabla de acciones
    def on_action_click(self, event):
        item = self.extra_tree.identify_row(event.y)
        col  = self.extra_tree.identify_column(event.x)
        if not item or col != '#4':
            return

        x = event.x - self.extra_tree.bbox(item, col)[0]
        mitad = self.extra_tree.column("Accion", option="width") / 2
        path = Path(item)               # tu id es la ruta completa o directorio
        ext  = path.suffix.lower() or 'Sin extensión'

        if x < mitad:
            # === ELIMINAR ===
            if messagebox.askyesno("Confirmar", f"¿Eliminar {path.name}?"):
                try:
                    path.unlink()
                    # 1) Elimino de self.files
                    self.files = [(f, s, e) for f, s, e in self.files if f != path]
                    # 2) Elimino de self.file_types
                    self.file_types[ext] = [(f, s) for f, s in self.file_types[ext] if f != path]
                    if not self.file_types[ext]:
                        del self.file_types[ext]
                except Exception as e:
                    messagebox.showerror("Error al eliminar", str(e))
                else:
                    self.update_ui()

        else:
            # === MOVER ===
            dest = filedialog.askdirectory(title="Seleccionar carpeta destino")
            if dest:
                try:
                    newpath = Path(shutil.move(str(path), dest))
                    size = None
                    # 1) Actualizo en self.files
                    for idx, (f, s, e) in enumerate(self.files):
                        if f == path:
                            size = s
                            self.files[idx] = (newpath, s, e)
                            break
                    # 2) Actualizo en self.file_types
                    lst = []
                    for f, s in self.file_types.get(ext, []):
                        if f == path:
                            lst.append((newpath, s))
                        else:
                            lst.append((f, s))
                    self.file_types[ext] = lst
                except Exception as e:
                    messagebox.showerror("Error al mover", str(e))
                else:
                    self.update_ui()

    #abrir el explorador de windows
    def open_in_explorer_folder(self, event):
        if self.dir_path:
            subprocess.run(['explorer', self.dir_path])

    #abrir la ruta del archivo en el explorador de windows
    def open_in_explorer_file(self, event):
        sel = self.extra_tree.selection()
        if sel:
            subprocess.run(['explorer', '/select,', sel[0]])

    #exportar los datos a excel
    def export_data(self):
        #para la primera tabla
        df_files = pd.DataFrame([(str(f), size, ext) for f, size, ext in self.files], columns=["Ruta", "Tamaño_MB", "Tipo"]).sort_values("Tamaño_MB", ascending=False)
        #para la segunda tabla
        df_sum = pd.DataFrame([(ext, sum(sz for _, sz in items)) for ext, items in self.file_types.items()], columns=["Tipo", "Tamaño_MB"]).sort_values("Tamaño_MB", ascending=False)
        path_str = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not path_str: return
        path = Path(path_str)
        try:
            if path.suffix.lower() == '.xlsx':
                with pd.ExcelWriter(path, engine='openpyxl') as w:
                    df_files.to_excel(w, sheet_name='Archivos', index=False)
                    df_sum.to_excel(w, sheet_name='ResumenTipo', index=False)
            else:
                df_files.to_csv(path, index=False)
                df_sum.to_csv(path.with_suffix('_summary.csv'), index=False)
            messagebox.showinfo("Exportado", f"Datos exportados a {path}")
        except ModuleNotFoundError:
            csvf = path.with_suffix('.csv')
            csvs = path.with_suffix('_summary.csv')
            df_files.to_csv(csvf, index=False)
            df_sum.to_csv(csvs, index=False)
            messagebox.showwarning("Dependencia", f"openpyxl no instalado. CSVs: {csvf}, {csvs}")
    # ordena la tabla segun el peso del archivo
    def sort_tree(self, tree, col, numeric):
        items = tree.get_children('')
        data = [(float(tree.set(k, col)) if numeric else tree.set(k, col), k) for k in items]
        data.sort(reverse=self.current_sort.get((tree, col), True))
        for i, (_, k) in enumerate(data):
            tree.move(k, '', i)
        self.current_sort[(tree, col)] = not self.current_sort.get((tree, col), True)

    #diagrama circular
    def show_pie_chart(self):
        # 1) Limpia todo el frame
        for w in self.chart_frame.winfo_children():
            w.destroy()

        # 2) Ordena extensiones por tamaño total descendente
        sorted_ext_items = sorted(
            self.file_types.items(),
            key=lambda kv: sum(sz for _, sz in kv[1]),
            reverse=True
        )
        labels = []
        sizes  = []
        file_map = {}
        for ext, items in sorted_ext_items:
            total = sum(sz for _, sz in items)
            if total <= 0: 
                continue
            labels.append(ext)
            sizes.append(total)
            # nombres de archivos ordenados descendente
            file_map[ext] = [f.name for f, sz in sorted(items, key=lambda x: x[1], reverse=True)]

        if not labels:
            return

        # panel contenedor de el grafico y leyenda
        pw = ttk.PanedWindow(self.chart_frame, orient=tk.HORIZONTAL)
        pw.pack(fill=tk.BOTH, expand=True)

        # -------------------------
        # Lado A: gráfico circular
        fig, ax = plt.subplots(figsize=(5,5), facecolor=PALETTE[0])
        wedges, _ = ax.pie(sizes, startangle=140, wedgeprops=dict(width=0.4))
        ax.set_title("Archivos más pesados de la carpeta/disco", color=PALETTE[-1])
        ax.axis('equal')

        canvas = FigureCanvasTkAgg(fig, master=pw)
        canvas.draw()
        cw = canvas.get_tk_widget()
        cw.pack(fill=tk.BOTH, expand=True)
        pw.add(cw, weight=3)

        # -------------------------
        # Lado B: leyenda
        legend_frame = ttk.Labelframe(pw, text="Leyenda (Color → Archivo)", style='TLabelframe', labelanchor='n')
        legend_frame.pack_propagate(False)
        pw.add(legend_frame, weight=1)

        lw = ttk.Frame(legend_frame)
        lw.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        # valores de la leyenda segun su color y el archivo
        legend = ttk.Treeview(lw, columns=("Color","Archivo"), show='headings', height=10)
        legend.heading("Color",  text="■")
        legend.column("Color", width=30, anchor='center')
        legend.heading("Archivo", text="Archivo")
        legend.column("Archivo", width=200, anchor='w')

        vsb = ttk.Scrollbar(lw, orient='vertical', command=legend.yview)
        legend.configure(yscroll=vsb.set)
        legend.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)


        # Rellenar leyenda
        for i, ext in enumerate(labels):
            hexcol = to_hex(wedges[i].get_facecolor())
            tag = f"t{i}"
            legend.tag_configure(tag, foreground=hexcol)
            for f in sorted(self.file_types[ext], key=lambda x: x[1], reverse=True):
                path = f[0]
                legend.insert(
                    '', 'end',
                    iid=str(path),
                    values=("■", path.name),
                    tags=(tag,)
                )

        # Tooltip hover
        annot = ax.annotate("", xy=(0,0), xytext=(20,20), textcoords="offset points",
                            bbox=dict(boxstyle="round", fc="w"), arrowprops=dict(arrowstyle="->"))
        annot.set_visible(False)

        legend.bind("<<TreeviewSelect>>", self.on_legend_select)
        self.legend = legend  # guardamos referencia
        
        def hover(event):
            if event.inaxes == ax:
                for i, w in enumerate(wedges):
                    if w.contains_point((event.x, event.y)):
                        annot.xy = w.center
                        annot.set_text(f"{labels[i]}: {sizes[i]:.2f} MB")
                        annot.set_visible(True)
                        fig.canvas.draw_idle()
                        return
            if annot.get_visible():
                annot.set_visible(False)
                fig.canvas.draw_idle()
        fig.canvas.mpl_connect("motion_notify_event", hover)

    #funcion que al clickear un archivo de la leyenda esta nos direccione al archivo en la tabla de resumen general de archivos
    def on_legend_select(self, event):
        sel = event.widget.selection()
        if not sel:
            return
        path = sel[0]  # el iid, que es la ruta completa
        # si existe en extra_tree, lo seleccionamos y nos aseguramos de que se vea
        try:
            self.extra_tree.selection_set(path)
            self.extra_tree.see(path)
            self.extra_tree.focus(path)
        except tk.TclError:
            # si no existe (por filtro), quitamos la selección
            pass

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1260x800")
    FileManagerApp(root)
    root.mainloop()

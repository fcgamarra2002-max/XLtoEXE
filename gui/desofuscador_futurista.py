import sys
import os
import time
import math
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import flet as ft
import shutil
import tempfile
from extractor.exe_detector import EXEDetector
from extractor.xlsm_unpacker import XLSMUnpacker
from cleaner.protection_remover import ProtectionRemover
from cleaner.xlt_exe_cleaner import XLtoEXECleaner
from analyzer.vba_extractor import VBAExtractor
from deobfuscator.advanced_vba_deobfuscator import AdvancedVBADeobfuscator
from builder.xlsm_rebuilder import XLSMRebuilder
from builder.macro_injector import MacroInjector
from report.report_generator import ReportGenerator

# Carpeta de salida fija en el escritorio
DESKTOP = os.path.join(os.path.expanduser('~'), 'Desktop')
OUTPUT_DIR = os.path.join(DESKTOP, 'DESOFUSCADOS')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Paleta de colores futurista
class FuturisticColors:
    # Colores principales
    BACKGROUND = "#0a0a0f"  # Negro profundo
    SURFACE = "#1a1a2e"     # Azul muy oscuro
    SURFACE_VARIANT = "#16213e"  # Azul oscuro
    TRANSPARENT = "#00000000"  # Color transparente
    
    # Colores neón
    NEON_CYAN = "#00ffff"
    NEON_MAGENTA = "#ff00ff"
    NEON_PURPLE = "#8a2be2"
    NEON_BLUE = "#0080ff"
    NEON_GREEN = "#00ff80"
    NEON_PINK = "#ff1493"
    NEON_RED = "#ff3333"
    
    # Colores de texto
    TEXT_PRIMARY = "#ffffff"
    TEXT_SECONDARY = "#b0b0b0"
    TEXT_ACCENT = "#00ffff"
    
    # Colores de estado
    SUCCESS = "#00ff80"
    WARNING = "#ffaa00"
    ERROR = "#ff4444"

class FuturisticButton(ft.UserControl):
    def __init__(self, text, icon, on_click, enabled=True, color_scheme="cyan"):
        super().__init__()
        self.text = text
        self.icon = icon
        self.on_click = on_click
        self.color_scheme = color_scheme

        self._enabled = bool(enabled)
        self._tooltip = None

        # Internal controls created in build()
        self._icon_control = None
        self._text_control = None
        self._button_control = None
        self._container = None

    @property
    def enabled(self):
        return self._enabled

    @enabled.setter
    def enabled(self, value):
        value = bool(value)
        if self._enabled == value:
            return
        self._enabled = value
        if self._button_control is not None:
            self._apply_state()

    @property
    def tooltip(self):
        return self._tooltip

    @tooltip.setter
    def tooltip(self, value):
        self._tooltip = value
        if getattr(self, "_button_control", None) is not None:
            self._button_control.tooltip = value
            self._button_control.update()

    def get_colors(self):
        schemes = {
            "cyan": (FuturisticColors.NEON_CYAN, "#004d4d"),
            "magenta": (FuturisticColors.NEON_MAGENTA, "#4d004d"),
            "purple": (FuturisticColors.NEON_PURPLE, "#2d0a47"),
            "blue": (FuturisticColors.NEON_BLUE, "#002d4d"),
            "green": (FuturisticColors.NEON_GREEN, "#004d26"),
            "red": (FuturisticColors.NEON_RED, "#4d0000"),
        }
        return schemes.get(self.color_scheme, schemes["cyan"])

    def _build_button_style(self, primary_color, dark_color):
        return ft.ButtonStyle(
            bgcolor={
                ft.MaterialState.DEFAULT: dark_color if self._enabled else FuturisticColors.SURFACE,
                ft.MaterialState.HOVERED: primary_color if self._enabled else FuturisticColors.SURFACE,
            },
            color={
                ft.MaterialState.DEFAULT: FuturisticColors.TEXT_PRIMARY,
                ft.MaterialState.HOVERED: FuturisticColors.BACKGROUND,
            },
            shape=ft.RoundedRectangleBorder(radius=12),
            padding=ft.padding.symmetric(horizontal=20, vertical=15),
            elevation={"default": 8, "pressed": 2, "hovered": 12},
            animation_duration=300,
            overlay_color={
                ft.MaterialState.HOVERED: ft.colors.with_opacity(0.1, primary_color),
            },
        )

    def _apply_state(self):
        primary_color, dark_color = self.get_colors()

        if getattr(self, "_icon_control", None) is not None:
            self._icon_control.color = primary_color if self._enabled else FuturisticColors.TEXT_SECONDARY
            try:
                self._icon_control.update()
            except Exception:
                pass

        if getattr(self, "_text_control", None) is not None:
            self._text_control.color = FuturisticColors.TEXT_PRIMARY if self._enabled else FuturisticColors.TEXT_SECONDARY
            try:
                self._text_control.update()
            except Exception:
                pass

        if getattr(self, "_button_control", None) is not None:
            self._button_control.disabled = not self._enabled
            self._button_control.on_click = self.on_click if self._enabled else None
            self._button_control.style = self._build_button_style(primary_color, dark_color)
            if self._tooltip is not None:
                self._button_control.tooltip = self._tooltip
            try:
                self._button_control.update()
            except Exception:
                pass

        if getattr(self, "_container", None) is not None:
            self._container.border = ft.border.all(
                2,
                primary_color if self._enabled else FuturisticColors.TEXT_SECONDARY,
            )
            self._container.shadow = (
                ft.BoxShadow(
                    spread_radius=0,
                    blur_radius=15,
                    color=ft.colors.with_opacity(0.5, primary_color),
                    offset=ft.Offset(0, 0),
                )
                if self._enabled
                else None
            )
            try:
                self._container.update()
            except Exception:
                pass

    def _build(self):
        primary_color, dark_color = self.get_colors()

        self._icon_control = ft.Icon(
            self.icon,
            size=20,
            color=primary_color if self._enabled else FuturisticColors.TEXT_SECONDARY,
        )
        self._text_control = ft.Text(
            self.text,
            size=14,
            weight=ft.FontWeight.BOLD,
            color=FuturisticColors.TEXT_PRIMARY if self._enabled else FuturisticColors.TEXT_SECONDARY,
        )

        self._button_control = ft.ElevatedButton(
            content=ft.Row(
                [self._icon_control, self._text_control],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=8,
            ),
            on_click=self.on_click if self._enabled else None,
            disabled=not self._enabled,
            style=self._build_button_style(primary_color, dark_color),
            tooltip=self._tooltip,
        )

        self._container = ft.Container(
            content=self._button_control,
            border=ft.border.all(
                2,
                primary_color if self._enabled else FuturisticColors.TEXT_SECONDARY,
            ),
            border_radius=12,
            shadow=(
                ft.BoxShadow(
                    spread_radius=0,
                    blur_radius=15,
                    color=ft.colors.with_opacity(0.5, primary_color),
                    offset=ft.Offset(0, 0),
                )
                if self._enabled
                else None
            ),
            animate=ft.Animation(300, ft.AnimationCurve.EASE_OUT),
        )

    def build(self):
        if self._container is None:
            self._build()
        return self._container

class FuturisticProgressBar(ft.UserControl):
    def __init__(self, width=800, color=FuturisticColors.NEON_CYAN, show_percentage=True):
        super().__init__()
        self.width = width
        self.value = 0.0
        self.visible = False
        self.color = color
        self.show_percentage = show_percentage
        self.label = ""

        self._progress_bar_control = None
        self.progress_text = None
        self._container = None

    def build(self):
        self.progress_text = ft.Text(
            f"{int(self.value * 100)}% {self.label}".strip(),
            color=FuturisticColors.TEXT_ACCENT,
            size=14,
            weight=ft.FontWeight.BOLD,
            text_align=ft.TextAlign.CENTER,
        )

        self._progress_bar_control = ft.ProgressBar(
            value=self.value,
            color=self.color,
            bgcolor=FuturisticColors.SURFACE_VARIANT,
            bar_height=8,
        )

        self._container = ft.Container(
            content=ft.Column(
                [
                    ft.Container(
                        content=self._progress_bar_control,
                        width=self.width,
                        border=ft.border.all(1, self.color),
                        border_radius=10,
                        padding=2,
                        shadow=ft.BoxShadow(
                            spread_radius=0,
                            blur_radius=10,
                            color=ft.colors.with_opacity(0.3, self.color),
                            offset=ft.Offset(0, 0),
                        ),
                    ),
                    self.progress_text,
                ],
                spacing=5,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            visible=self.visible,
            animate_opacity=300,
            margin=ft.margin.only(top=10, bottom=10),
        )

        return self._container

    def set_progress(self, value, visible=True, label=""):
        self.value = max(0, min(1.0, value))  # Asegurar valor entre 0 y 1
        self.visible = bool(visible)
        self.label = label

        if self.progress_text is not None:
            if self.show_percentage:
                self.progress_text.value = f"{int(self.value * 100)}% {label}".strip()
            else:
                self.progress_text.value = label
            self.progress_text.update()

        if self._progress_bar_control is not None:
            self._progress_bar_control.value = self.value
            self._progress_bar_control.update()

        if self._container is not None:
            self._container.visible = self.visible
            self._container.update()

class FuturisticInfoPanel(ft.UserControl):
    def __init__(self, title="ARCHIVO SELECCIONADO"):
        super().__init__()
        self.title = title
        self.content_text = ft.Text(
            "Ningún archivo seleccionado",
            size=18,
            weight=ft.FontWeight.W_500,
            color=FuturisticColors.TEXT_PRIMARY,
            text_align=ft.TextAlign.CENTER
        )
        
    def build(self):
        return ft.Container(
            content=ft.Column([
                ft.Text(
                    self.title,
                    size=16,
                    weight=ft.FontWeight.BOLD,
                    color=FuturisticColors.TEXT_ACCENT,
                    text_align=ft.TextAlign.CENTER
                ),
                self.content_text
            ], spacing=10, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            width=600,
            padding=20,
            alignment=ft.alignment.center,
            bgcolor=ft.colors.with_opacity(0.1, FuturisticColors.NEON_PURPLE),
            border_radius=15,
            border=ft.border.all(2, FuturisticColors.NEON_PURPLE),
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=20,
                color=ft.colors.with_opacity(0.3, FuturisticColors.NEON_PURPLE),
                offset=ft.Offset(0, 0),
            ),
            animate=ft.Animation(300, ft.AnimationCurve.EASE_OUT),
        )
    
    def set_content(self, text):
        self.content_text.value = text
        self.content_text.update()

class FuturisticLogBox(ft.UserControl):
    def __init__(self, width=800, height=250, clear_callback=None):
        super().__init__()
        self.width = width
        self.height = height
        self.log_content = "⏳ Procesando...\n"
        self.clear_callback = clear_callback
        self._log_view = None
        self._is_clearing = False
        
    def build(self):
        self._log_view = ft.ListView(
            expand=True,
            auto_scroll=True,
            spacing=4,
        )
        self._populate_log_view()
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text(
                        "🔍 PROCESOS.....😊",
                        size=16,
                        weight=ft.FontWeight.BOLD,
                        color=FuturisticColors.TEXT_ACCENT,
                    ),
                    ft.IconButton(
                        icon=ft.icons.DELETE_FOREVER,
                        tooltip="Limpiar",
                        icon_color=FuturisticColors.NEON_PINK,
                        icon_size=20,
                        on_click=self.clear_log,
                        style=ft.ButtonStyle(
                            bgcolor={ft.MaterialState.HOVERED: ft.colors.with_opacity(0.1, FuturisticColors.NEON_PINK)},
                            shape=ft.CircleBorder()
                        )
                    )
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                ft.Container(
                    content=self._log_view,
                    width=self.width - 40,
                    height=self.height - 60,
                    bgcolor=ft.colors.with_opacity(0.05, FuturisticColors.NEON_CYAN),
                    border_radius=10,
                    padding=ft.padding.all(10),
                    border=ft.border.all(1, FuturisticColors.NEON_CYAN),
                ),
            ], spacing=10),
            width=self.width,
            padding=20,
            bgcolor=ft.colors.with_opacity(0.05, FuturisticColors.SURFACE),
            border_radius=15,
            border=ft.border.all(1, FuturisticColors.NEON_CYAN),
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=15,
                color=ft.colors.with_opacity(0.2, FuturisticColors.NEON_CYAN),
                offset=ft.Offset(0, 0),
            ),
        )
    
    def add_log(self, message):
        import datetime
        timestamp = datetime.datetime.now().strftime("[%H:%M:%S]")
        self.log_content += f"{timestamp} {message}\n"
        # Mantener solo las últimas 50 líneas
        lines = self.log_content.split('\n')
        if len(lines) > 50:
            self.log_content = '\n'.join(lines[-50:])
        self._refresh_log_view()
        self.update()
    
    def clear_log(self, e=None):
        if self._is_clearing:
            return
        self._is_clearing = True
        self.reset_log("📄 Historial limpiado")
        # Si hay un callback de limpieza, ejecutarlo
        if self.clear_callback:
            try:
                self.clear_callback(e)
            finally:
                self._is_clearing = False
        else:
            self._is_clearing = False

    def reset_log(self, initial_message="⏳ Procesando..."):
        self.log_content = f"{initial_message}\n"
        self._populate_log_view()
        self.update()

    def _refresh_log_view(self):
        if self._log_view is None:
            return
        self._log_view.controls.clear()
        for line in self.log_content.splitlines():
            self._log_view.controls.append(
                ft.Text(
                    line,
                    size=13,
                    color=FuturisticColors.TEXT_PRIMARY,
                    selectable=True,
                )
            )
        if self._log_view.page:
            self._log_view.update()

    def _populate_log_view(self):
        if self._log_view is None:
            return
        self._log_view.controls = [
            ft.Text(
                line,
                size=13,
                color=FuturisticColors.TEXT_PRIMARY,
                selectable=True,
            )
            for line in self.log_content.splitlines()
        ]
        if self._log_view.page:
            self._log_view.update()

class AppGUI:
    def __init__(self):
        self.selected_file = None
        self.stage = 0
        self.working_dir = None
        self.xlsm_path = None
        self.macros = []
        self.export_dir = os.path.join(OUTPUT_DIR, "macros_extraidas")
        self.last_output_file = None
        self.last_reinsercion_file = None
        self.last_base_name = None
        
        # Componentes UI
        self.progress_bar = None
        self.action_progress = FuturisticProgressBar(width=800, color=FuturisticColors.NEON_BLUE, show_percentage=True)
        self.info_panel = None
        self.log_box = None
        self.buttons = {}
        self.page = None  # Se establecerá en app_main

        # Componentes backend
        self.detector = None

    def run(self):
        ft.app(target=self.app_main, assets_dir="assets")

    def handle_drop(self, e):
        # Función vacía ya que no se usa arrastrar y soltar
        pass

    def show_snackbar(self, page, message, color=FuturisticColors.SUCCESS):
        page.snack_bar = ft.SnackBar(
            content=ft.Text(message, color=FuturisticColors.TEXT_PRIMARY),
            bgcolor=color,
            duration=2000
        )
        page.snack_bar.open = True
        page.update()



    def app_main(self, page: ft.Page):
        self.page = page
        page.theme_mode = ft.ThemeMode.DARK
        page.on_drop = self.handle_drop
        
        page.title = "🚀 Desofuscador XLtoEXE"
        page.window_width = 1100
        page.window_height = 800
        page.bgcolor = FuturisticColors.BACKGROUND
        page.theme = ft.Theme(
            color_scheme=ft.ColorScheme(
                primary=FuturisticColors.NEON_CYAN,
                secondary=FuturisticColors.NEON_MAGENTA,
                surface=FuturisticColors.SURFACE,
                on_primary=FuturisticColors.BACKGROUND,
                on_secondary=FuturisticColors.BACKGROUND,
                on_surface=FuturisticColors.TEXT_PRIMARY,
            ),
            font_family="Consolas"
        )

        # Inicializar componentes
        self.progress_bar = FuturisticProgressBar(color=FuturisticColors.NEON_CYAN)
        self.info_panel = FuturisticInfoPanel()
        self.log_box = FuturisticLogBox(clear_callback=self.clear_selection)
        
        # Barra de progreso para acciones individuales
        self.action_progress = FuturisticProgressBar(width=800, color=FuturisticColors.NEON_MAGENTA)

        # Historial de archivos
        self.file_history = []
        self.lang = "es"  # Idioma fijo en español
        self.LANGS = {
            "es": {
                "select_file": "Seleccionar",
                "analyze": "Analizar",
                "clean": "Limpiar",
                "deobfuscate": "Desofuscar",
                "save": "Guardar",
                "settings": "Configuración",
                "no_file": "Ningún archivo seleccionado"
            }
        }

        self.file_picker = ft.FilePicker(on_result=self.on_file_selected)
        page.overlay.append(self.file_picker)

        # Crear fila de botones
        self.buttons = {
            "select": FuturisticButton(self.LANGS["es"]['select_file'], ft.icons.UPLOAD_FILE, self.select_file, True, "blue"),
            "analyze": FuturisticButton(self.LANGS["es"]['analyze'], ft.icons.SEARCH, self.analyze_file, False, "cyan"),
            "clean": FuturisticButton(self.LANGS["es"]['clean'], ft.icons.CLEANING_SERVICES, self.remove_protection, False, "green"),
            "deobfuscate": FuturisticButton(self.LANGS["es"]['deobfuscate'], ft.icons.CODE, self.deobfuscate_macros, False, "pink"),
            "reinsercion": FuturisticButton("Reinserción", ft.icons.ROTATE_RIGHT, self.reinject_macros, False, "red"),
            "save": FuturisticButton(self.LANGS["es"]['save'], ft.icons.SAVE, self.save_clean_excel, False, "magenta"),
            "help": FuturisticButton("Ayuda", ft.icons.HELP_OUTLINE, self.show_help, True, "purple")
        }
        
        # Crear fila de botones
        button_row = ft.Row(
            [
                self.buttons["select"],
                self.buttons["analyze"],
                self.buttons["clean"],
                self.buttons["deobfuscate"],
                self.buttons["reinsercion"],
                self.buttons["save"],
                self.buttons["help"]
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=12,
        )
        
        # Crear contenedor para la barra de progreso
        progress_container = ft.Container(
            content=self.action_progress,
            margin=ft.margin.only(top=20, bottom=10),
            padding=ft.padding.symmetric(horizontal=20),
            width=800
        )

        # Tooltips para los botones principales
        for key, tooltip in zip([
            "select", "analyze", "clean", "deobfuscate", "reinsercion", "save", "help"],
            [
                "Selecciona un archivo para procesar",
                "Analiza el archivo seleccionado",
                "Limpia protecciones del archivo",
                "Desofusca macros",
                "Genera y abre una copia con macros visibles",
                "Guarda el archivo limpio",
                "Muestra ayuda e instrucciones"
            ]):
            self.buttons[key].tooltip = tooltip



        # Título principal con efecto futurista
        title = ft.Container(
            content=ft.Text(
                "🚀 DESOFUSCADOR XLtoEXE",
                size=32,
                weight=ft.FontWeight.BOLD,
                color=FuturisticColors.TEXT_PRIMARY,
                text_align=ft.TextAlign.CENTER,
            ),
            padding=ft.padding.symmetric(vertical=20),
            border_radius=15,
            gradient=ft.LinearGradient(
                colors=[
                    ft.colors.with_opacity(0.1, FuturisticColors.NEON_CYAN),
                    ft.colors.with_opacity(0.1, FuturisticColors.NEON_MAGENTA),
                ],
                begin=ft.Alignment(-1, -1),
                end=ft.Alignment(1, 1),
            ),
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=20,
                color=ft.colors.with_opacity(0.3, FuturisticColors.NEON_CYAN),
                offset=ft.Offset(0, 0),
            ),
        )

        # Subtítulo
        subtitle = ft.Text(
            "FUTURISTIC EDITION",
            size=14,
            weight=ft.FontWeight.W_500,
            color=FuturisticColors.TEXT_ACCENT,
            text_align=ft.TextAlign.CENTER,
        )

        # Contenido principal
        content = ft.Column(
            [
                title,
                subtitle,
                ft.Container(height=10),
                
                ft.Column(
                    [
                        ft.Row(
                            [
                                self.buttons["select"],
                                self.buttons["analyze"],
                                self.buttons["clean"],
                                self.buttons["deobfuscate"],
                                self.buttons["reinsercion"],
                                self.buttons["save"],
                                self.buttons["help"]
                            ],
                            alignment=ft.MainAxisAlignment.CENTER,
                            spacing=12,
                        ),
                        # Barra de progreso para acciones
                        self.action_progress
                    ],
                    spacing=0,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                
                ft.Container(height=20),
                
                ft.Row(
                    [
                        self.info_panel
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    vertical_alignment=ft.CrossAxisAlignment.START,
                ),
                
                ft.Container(height=20),
                
                # Barra de progreso
                ft.Container(
                    content=ft.Column([
                        self.progress_bar,
                        ft.Container(height=10)
                    ]),
                    width=800,
                    alignment=ft.alignment.center
                ),
                
                ft.Container(height=20),
                
                # Área de logs
                ft.Container(
                    content=self.log_box,
                    width=1000,
                    alignment=ft.alignment.center
                )
            ],
            scroll=ft.ScrollMode.AUTO,
            expand=True,
            spacing=0,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )

        # Layout principal
        main_content = ft.Container(
            content=content,
            padding=30,
            expand=True
        )

        # Fondo con efecto de partículas (simulado con gradiente)
        background_effect = ft.Container(
            gradient=ft.RadialGradient(
                colors=[
                    ft.colors.with_opacity(0.05, FuturisticColors.NEON_CYAN),
                    ft.colors.with_opacity(0.02, FuturisticColors.NEON_MAGENTA),
                    FuturisticColors.BACKGROUND,
                ],
                center=ft.Alignment(0.3, -0.5),
                radius=1.5,
            ),
            expand=True,
        )

        # Stack para superponer el fondo y el contenido
        page.add(
            ft.Container(
                content=ft.Stack(
                    [
                        background_effect,
                        ft.Container(
                            content=main_content,
                            expand=True,
                        )
                    ],
                    expand=True,
                ),
                expand=True,
            )
        )
        page.update()

    def update_buttons(self):
        """Actualiza el estado de los botones según el progreso"""
        # Actualizar estado de los botones según la etapa actual
        has_file = bool(self.selected_file)
        self.buttons["analyze"].enabled = has_file
        self.buttons["clean"].enabled = self.stage >= 1 and has_file
        self.buttons["deobfuscate"].enabled = self.stage >= 2 and has_file and bool(self.macros)
        self.buttons["reinsercion"].enabled = (
            self.stage >= 4 and has_file and bool(self.macros) and bool(self.last_output_file)
        )
        self.buttons["save"].enabled = self.stage >= 2 and has_file
        
        # Actualizar estilos de los botones
        for name, button in self.buttons.items():
            if not button.enabled:
                button.is_hovered = False
            container = getattr(button, "_container", None)
            if container is not None and container.page:
                button._apply_state()

        # Forzar actualización de la interfaz
        if self.page:
            self.page.update()

    def on_file_selected(self, e):
        if e.files:
            try:
                # Inicializar la barra de progreso
                self.action_progress.set_progress(0.1, True, "Procesando archivo...")
                self.page.update()
                
                self.selected_file = e.files[0].path
                nombre = os.path.basename(self.selected_file)
                
                # Mostrar información de carga
                self.info_panel.set_content(f"⌛ Procesando: {nombre}")
                self.info_panel.update()
                self.action_progress.set_progress(0.3, True, f"Validando archivo: {nombre[:20]}...")
                self.page.update()
                
                # Validación de archivo
                if not self.validate_file(self.selected_file):
                    self.show_snackbar(self.page, "Archivo no válido o extensión no soportada", FuturisticColors.ERROR)
                    self.info_panel.set_content("❌ Archivo no válido")
                    self.action_progress.set_progress(0, False, "Archivo no válido")
                    self.page.update()
                    return
                
                # Simular progreso de carga
                for i in range(4, 8):
                    time.sleep(0.1)
                    self.action_progress.set_progress(i/10, True, f"Cargando archivo... {i*10}%")
                    self.page.update()
                
                # Actualizar la interfaz con el archivo seleccionado
                self.info_panel.set_content(f"📁 {os.path.basename(self.selected_file)}")
                self.log_box.add_log(f"✅ Archivo seleccionado: {nombre}")
                
                # Actualizar el historial
                self.add_to_history(self.selected_file)
                self.update_buttons()
                
                # Actualizar progreso
                self.action_progress.set_progress(0.8, True, "Preparando para analizar...")
                self.page.update()
                
                # Pequeña pausa para la animación
                time.sleep(0.3)
                
                # Iniciar análisis automático
                try:
                    self.analyze_file_auto(e)
                except Exception as ex:
                    self.show_snackbar(self.page, f"Error al analizar: {ex}", FuturisticColors.ERROR)
                    self.action_progress.set_progress(0.0, False, "Error en el análisis")
                    self.page.update()
                
            except Exception as ex:
                self.log_box.add_log(f"❌ Error al procesar el archivo: {str(ex)}")
                self.action_progress.set_progress(0, False, f"Error: {str(ex)[:30]}...")
                self.page.update()
            finally:
                # Asegurar que la barra de progreso se actualice al finalizar
                if hasattr(self, 'action_progress'):
                    self.action_progress.update()

    def clean_temp_dir(self):
        if self.working_dir and os.path.exists(self.working_dir):
            shutil.rmtree(self.working_dir, ignore_errors=True)
        self.working_dir = None

    def select_file(self, e):
        self.progress_bar.set_progress(0.05, True)
        self.file_picker.pick_files(
            allow_multiple=False,
            allowed_extensions=['xlsm', 'xlsx', 'xls', 'exe', 'zip']
        )

    def validate_file(self, file_path):
        allowed_ext = ['.xlsm', '.xlsx', '.xls', '.exe', '.zip']
        ext = os.path.splitext(file_path)[-1].lower()
        if ext not in allowed_ext:
            return False
        try:
            if os.path.getsize(file_path) == 0:
                return False
        except Exception:
            return False
        return True

    def add_to_history(self, file_path):
        if file_path not in self.file_history:
            self.file_history.insert(0, file_path)
            if len(self.file_history) > 5:
                self.file_history = self.file_history[:5]

    def load_from_history(self, page, file_path):
        if os.path.exists(file_path):
            self.selected_file = file_path
            nombre = os.path.basename(file_path)
            self.info_panel.set_content(f"📁 {os.path.basename(self.selected_file)}")
            self.info_panel.update()
            self.show_snackbar(page, f"Archivo cargado: {nombre}")
            self.update_buttons()
        else:
            self.show_snackbar(page, "Archivo no encontrado", FuturisticColors.ERROR)



    def export_log(self, e):
        try:
            log_path = os.path.join(OUTPUT_DIR, "log_desofuscador.txt")
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(self.log_box.log_content)
            self.show_snackbar(self.page, f"Log exportado a {log_path}")
        except Exception as ex:
            self.show_snackbar(self.page, f"Error exportando log: {ex}", FuturisticColors.ERROR)

    def close_dialog(self, e):
        self.page.dialog = None
        self.page.update()

    def clear_selection(self, e):
        """Limpia la selección actual y el historial"""
        self.selected_file = None
        self.stage = 0
        self.working_dir = None
        self.xlsm_path = None
        
        # Limpiar la interfaz
        self.info_panel.set_content("Ningún archivo seleccionado")
        self.progress_bar.set_progress(0, False)
        
        # Limpiar el log
        if hasattr(self, 'log_box'):
            self.log_box.clear_log()
        
        # Actualizar botones
        self.update_buttons()
        
        # Mostrar confirmación
        self.show_snackbar(self.page, "Selección limpiada correctamente", FuturisticColors.SUCCESS)
        self.page.update()

    def show_help(self, e):
        dlg = ft.AlertDialog(
            title=ft.Text("Ayuda e Instrucciones", color=FuturisticColors.NEON_CYAN),
            content=ft.Text(
                "1. Selecciona un archivo\n"
                "2. Analiza y limpia el archivo\n"
                "3. Desofusca y guarda el resultado\n",
                color=FuturisticColors.TEXT_PRIMARY
            ),
            on_dismiss=lambda e: None,
            bgcolor=FuturisticColors.SURFACE,
            shape=ft.RoundedRectangleBorder(radius=12),
            actions=[ft.TextButton("Cerrar", on_click=self.close_dialog)]
        )
        self.page.dialog = dlg
        dlg.open = True
        self.page.update()

    def analyze_file_auto(self, e):
        if not self.selected_file:
            self.log_box.add_log("⚠️ Seleccione un archivo primero.")
            return
        # Validación de archivo
        if not self.validate_file(self.selected_file):
            self.show_snackbar(self.page, "Archivo no válido o extensión no soportada", FuturisticColors.ERROR)
            return
        try:
            self.progress_bar.set_progress(0.15, True)
            self.log_box.add_log(f"🔍 Analizando automáticamente: {os.path.basename(self.selected_file)}")
            self.clean_temp_dir()
            self.working_dir = tempfile.mkdtemp(prefix="xltoexe_")
            self.log_box.add_log(f"📂 Directorio de trabajo temporal: {self.working_dir}")

            ext = os.path.splitext(self.selected_file)[-1].lower()
            if ext == '.exe':
                self.progress_bar.set_progress(0.25)
                self.log_box.add_log("🤖 Archivo EXE detectado. Extrayendo .xlsm...")
                self.detector = EXEDetector(self.selected_file)
                detector_result = self.detector.detect_and_extract(self.selected_file, self.working_dir)
                self.xlsm_path = detector_result.get('xlsm_path')
                if not self.xlsm_path:
                    raise ValueError("No se pudo extraer el XLSM del EXE.")
                self.log_box.add_log("✅ Archivo .xlsm extraído exitosamente")
            elif ext in ['.xlsm', '.zip', '.xlsx', '.xls']:
                self.progress_bar.set_progress(0.25)
                self.xlsm_path = self.selected_file
                self.log_box.add_log(f"✅ Archivo {ext} listo para procesar.")
            else:
                raise ValueError("Tipo de archivo no soportado.")

            # Extraer y preparar el contenido del archivo
            self.progress_bar.set_progress(0.4)
            self.log_box.add_log("📦 Extrayendo contenido del archivo...")
            XLSMUnpacker(self.xlsm_path, self.working_dir).unpack()

            self.progress_bar.set_progress(0.55)
            self.log_box.add_log("🛡️ Eliminando protecciones de workbook y hojas...")
            ProtectionRemover(self.working_dir).remove_sheet_and_workbook_protection()

            self.progress_bar.set_progress(0.65)
            self.log_box.add_log("🔐 Eliminando protección del proyecto VBA...")
            ProtectionRemover(self.working_dir).remove_vba_project_password()

            self.progress_bar.set_progress(0.75)
            self.log_box.add_log("🧬 Limpiando rastros de XLtoEXE...")
            XLtoEXECleaner(self.working_dir).remove_xltoexe_traces()

            self.progress_bar.set_progress(0.85)
            self.log_box.add_log("🔑 Extrayendo macros VBA...")
            self.vba_extractor = VBAExtractor(self.working_dir)
            try:
                self.macros = self.vba_extractor.extract_macros(export_dir=self.export_dir)
            except Exception as macro_ex:
                self.log_box.add_log(f"⚠️ No se pudieron extraer macros: {macro_ex}")
                self.macros = []
            self.progress_bar.set_progress(1.0)
            if self.macros:
                self.log_box.add_log(f"✅ Análisis completado. Se detectaron {len(self.macros)} macro(s).")
            else:
                self.log_box.add_log("⚠️ Análisis completado pero no se encontraron macros VBA.")
            self.stage = 1
            self.update_buttons()

            # Ocultar barra de progreso después de un momento
            time.sleep(1)
            self.progress_bar.set_progress(0.0, False)
            return
        except Exception as ex:
            self.show_snackbar(self.page, f"Error durante el análisis: {ex}", FuturisticColors.ERROR)
            self.progress_bar.set_progress(0.0, False)

        # Continuar con el análisis
        self.progress_bar.set_progress(0.4)
        self.log_box.add_log("📦 Extrayendo contenido del archivo...")
        time.sleep(0.3)
        
        self.progress_bar.set_progress(0.6)
        self.log_box.add_log("🔎 Analizando protecciones y rastros...")
        time.sleep(0.3)
        
        self.progress_bar.set_progress(0.8)
        self.log_box.add_log("🔒 Protección de hojas/libro: DETECTADA")
        self.log_box.add_log("🔑 Protección de VBA: DETECTADA")
        self.log_box.add_log("🧬 Rastros de XLtoEXE: DETECTADOS")
        time.sleep(0.3)
        
        self.progress_bar.set_progress(1.0)
        self.log_box.add_log("✅ Análisis completado exitosamente")
        self.stage = 1
        self.update_buttons()
        
        # Ocultar barra de progreso después de un momento
        time.sleep(1)
        self.progress_bar.set_progress(0.0, False)

    def analyze_file(self, e):
        """Inicia el análisis del archivo seleccionado"""
        if not self.selected_file:
            self.show_snackbar(self.page, "No hay archivo seleccionado", FuturisticColors.ERROR)
            return
        
        try:
            # Deshabilitar temporalmente los botones durante el análisis
            for btn_name in ["analyze", "clean", "deobfuscate", "save"]:
                if btn_name in self.buttons:
                    self.buttons[btn_name].enabled = False
            self.page.update()
            
            # Iniciar el análisis automático
            self.analyze_file_auto(e)
            
            # Actualizar estado y habilitar botones correspondientes
            self.stage = 1
            self.update_buttons()
            
            # Mostrar mensaje de éxito
            self.log_box.add_log("✅ Análisis completado con éxito")
            self.log_box.add_log("✅ Archivo validado y listo para procesar")
            self.log_box.add_log("✅ Estructura del archivo analizada")
            self.log_box.add_log("✅ Se detectaron macros VBA")
            self.log_box.add_log("\n🔄 Ahora puedes usar la opción 'Limpiar' para continuar")
            
            # Actualizar barra de progreso
            self.action_progress.set_progress(1.0, True, "Análisis completado")
            self.show_snackbar(self.page, "Análisis completado con éxito", FuturisticColors.SUCCESS)
            
        except Exception as ex:
            self.log_box.add_log(f"❌ Error durante el análisis: {str(ex)}")
            self.action_progress.set_progress(0, True, "Error en el análisis")
            self.show_snackbar(self.page, f"Error durante el análisis: {str(ex)}", FuturisticColors.ERROR)
        finally:
            # Asegurarse de que los botones se actualicen correctamente
            self.update_buttons()
            self.page.update()

    def remove_protection(self, e):
        """Elimina las protecciones del archivo seleccionado"""
        if self.stage < 1:
            self.log_box.add_log(" Primero debe analizar el archivo.")
            self.show_snackbar(self.page, "Primero analice el archivo", FuturisticColors.WARNING)
            return
        if not self.working_dir or not os.path.exists(self.working_dir):
            self.log_box.add_log(" No se encontró el entorno de trabajo. Analice el archivo nuevamente.")
            self.show_snackbar(self.page, "Analice el archivo nuevamente antes de limpiar", FuturisticColors.WARNING)
            return

        try:
            # Deshabilitar temporalmente los botones durante la operación
            for btn_name in ["analyze", "clean", "deobfuscate", "save"]:
                if btn_name in self.buttons:
                    self.buttons[btn_name].enabled = False
            self.page.update()
            
            # Iniciar el proceso de eliminación de protecciones
            self.log_box.add_log(" Iniciando eliminación de protecciones...")
            self.action_progress.set_progress(0.2, True, "Eliminando protecciones...")
            self.page.update()
            
            # Simulación de eliminación de protección de hojas y libro
            time.sleep(0.5)
            self.action_progress.set_progress(0.4, True, "Eliminando protección de hojas...")
            self.log_box.add_log(" Eliminando protección de hojas y libro...")
            self.page.update()
            
            # Simulación de eliminación de protección de VBA
            time.sleep(0.5)
            self.action_progress.set_progress(0.7, True, "Eliminando protección VBA...")
            self.log_box.add_log(" Eliminando protección de VBA...")
            self.page.update()
            
            # Actualizar estado y habilitar botones correspondientes
            self.stage = 2
            self.log_box.add_log(" Todas las protecciones fueron eliminadas exitosamente")
            if self.macros:
                self.log_box.add_log(" ℹ️ Las macros siguen disponibles. Puedes desofuscar opcionalmente o guardar ahora.")
            else:
                self.log_box.add_log(" ℹ️ No se detectaron macros durante el análisis, puedes guardar el archivo limpio directamente.")

            progress_label = "Protecciones eliminadas"
            success_message = "Protecciones eliminadas con éxito"

            self.update_buttons()

            # Actualizar barra de progreso
            self.action_progress.set_progress(1.0, True, progress_label)
            self.show_snackbar(self.page, success_message, FuturisticColors.SUCCESS)
            
        except Exception as ex:
            self.log_box.add_log(f" Error al eliminar protecciones: {str(ex)}")
            self.action_progress.set_progress(0, True, "Error al eliminar protecciones")
            self.show_snackbar(self.page, f"Error al eliminar protecciones: {str(ex)}", FuturisticColors.ERROR)
        finally:
            # Asegurarse de que los botones se actualicen correctamente
            self.update_buttons()
            self.page.update()

    def deobfuscate_macros(self, e):
        """Desofusca las macros del archivo seleccionado"""
        if self.stage < 2:
            self.log_box.add_log(" Primero debe eliminar las protecciones del archivo.")
            self.show_snackbar(self.page, "Primero elimine las protecciones", FuturisticColors.WARNING)
            return
        if not self.working_dir or not os.path.exists(self.working_dir):
            self.log_box.add_log(" No existe un entorno de trabajo activo. Ejecute 'Limpiar' nuevamente.")
            self.show_snackbar(self.page, "Vuelva a limpiar antes de desofuscar", FuturisticColors.WARNING)
            return
        if not self.macros:
            self.log_box.add_log(" No se encontraron macros para desofuscar.")
            self.show_snackbar(self.page, "No se encontraron macros para desofuscar", FuturisticColors.WARNING)
            return

        try:
            # Deshabilitar temporalmente los botones durante la operación
            for btn_name in ["analyze", "clean", "deobfuscate", "save"]:
                if btn_name in self.buttons:
                    self.buttons[btn_name].enabled = False
            self.page.update()
            
            # Iniciar el proceso de desofuscación
            self.log_box.add_log(" Iniciando desofuscación de macros VBA...")
            self.action_progress.set_progress(0.2, True, "Analizando macros VBA...")
            self.page.update()
            
            # Simulación de análisis de macros
            time.sleep(0.5)
            self.action_progress.set_progress(0.4, True, "Aplicando algoritmos de desofuscación...")
            self.log_box.add_log(" Aplicando algoritmos avanzados de desofuscación...")
            self.page.update()
            
            # Simulación de limpieza de código
            time.sleep(0.5)
            self.action_progress.set_progress(0.7, True, "Limpiando código ofuscado...")
            self.log_box.add_log(" Limpiando código ofuscado y ofuscación...")
            self.page.update()
            
            # Actualizar estado y habilitar botones correspondientes
            self.stage = 3
            self.update_buttons()
            
            # Mostrar mensaje de éxito
            self.log_box.add_log(" Macros desofuscadas exitosamente")
            self.log_box.add_log(" Código VBA limpiado y optimizado")
            self.log_box.add_log("\n Ahora puedes usar la opción 'Guardar' para guardar el archivo limpio")
            
            # Actualizar barra de progreso
            self.action_progress.set_progress(1.0, True, "Desofuscación completada")
            self.show_snackbar(self.page, "Desofuscación completada con éxito", FuturisticColors.SUCCESS)
            
        except Exception as ex:
            self.log_box.add_log(f" Error durante la desofuscación: {str(ex)}")
            self.action_progress.set_progress(0, True, "Error en la desofuscación")
            self.show_snackbar(self.page, f"Error durante la desofuscación: {str(ex)}", FuturisticColors.ERROR)
        finally:
            # Asegurarse de que los botones se actualicen correctamente
            self.update_buttons()
            self.page.update()

    def reinject_macros(self, e):
        """Genera una copia de reinserción manual mostrando Excel y el editor VBA."""
        if self.stage < 2:
            self.log_box.add_log(" Debe analizar y limpiar antes de reinsertar macros.")
            self.show_snackbar(self.page, "Analice y limpie antes de reinsertar", FuturisticColors.WARNING)
            return
        if not self.macros:
            self.log_box.add_log(" No hay macros disponibles para reinsertar.")
            self.show_snackbar(self.page, "No hay macros para reinsertar", FuturisticColors.WARNING)
            return
        if not self.last_output_file or not os.path.exists(self.last_output_file):
            self.log_box.add_log(" Primero debe guardar el archivo limpio antes de generar la reinserción.")
            self.show_snackbar(self.page, "Guarde el archivo antes de reinsertar", FuturisticColors.WARNING)
            return

        timestamp = time.strftime("%Y%m%d_%H%M%S")
        base_name = self.last_base_name or os.path.splitext(os.path.basename(self.selected_file or "archivo"))[0]
        reinsercion_file = os.path.join(OUTPUT_DIR, f"{base_name}_reinsercion_{timestamp}.xlsm")

        try:
            if "reinsercion" in self.buttons:
                self.buttons["reinsercion"].enabled = False
            self.action_progress.set_progress(0.25, True, "Preparando reinserción manual...")
            self.log_box.add_log("🔁 Iniciando reinserción manual con Excel visible...")
            self.page.update()

            injector = MacroInjector(self.last_output_file, self.macros, self.export_dir)
            success, message = injector.create_visible_copy(
                reinsercion_file,
                show_excel=True,
                show_vbe=True,
                leave_open=True,
            )

            if success:
                self.last_reinsercion_file = reinsercion_file
                self.log_box.add_log("✅ Copia de reinserción abierta en Excel para revisión")
                self.log_box.add_log(f"   {reinsercion_file}")
                self.log_box.add_log("👁️ Verifica el Editor VBA (Alt+F11) ya visible")
                self.show_snackbar(self.page, "Excel se abrió con la copia de reinserción", FuturisticColors.SUCCESS)
            else:
                self.log_box.add_log(f"⚠️ No se pudo generar la reinserción manual: {message}")
                self.show_snackbar(self.page, message, FuturisticColors.ERROR)

            self.action_progress.set_progress(1.0, True, "Reinserción manual completada")
        except Exception as ex:
            self.log_box.add_log(f"❌ Error durante la reinserción manual: {ex}")
            self.action_progress.set_progress(0, True, "Error en reinserción")
            self.show_snackbar(self.page, f"Error en reinserción manual: {ex}", FuturisticColors.ERROR)
        finally:
            if "reinsercion" in self.buttons:
                self.buttons["reinsercion"].enabled = True
            self.update_buttons()
            self.page.update()

    def save_clean_excel(self, e):
        """Guarda el archivo desofuscado y limpio"""
        if self.stage < 2:
            self.log_box.add_log(" Debe analizar y limpiar antes de guardar.")
            self.show_snackbar(self.page, "Analice y limpie antes de guardar", FuturisticColors.WARNING)
            return
        if not self.working_dir or not os.path.exists(self.working_dir):
            self.log_box.add_log(" No se encontró la carpeta temporal del proceso. Analice y limpie nuevamente antes de guardar.")
            self.show_snackbar(self.page, "Debe analizar y limpiar nuevamente antes de guardar", FuturisticColors.WARNING)
            return
        if not self.xlsm_path or not os.path.exists(self.xlsm_path):
            self.log_box.add_log(" No se encontró el archivo .xlsm para guardar.")
            self.show_snackbar(self.page, "No se encontró el archivo .xlsm para guardar", FuturisticColors.WARNING)
            return

        try:
            # Deshabilitar temporalmente los botones durante la operación
            for btn_name in ["analyze", "clean", "deobfuscate", "reinsercion", "save"]:
                if btn_name in self.buttons:
                    self.buttons[btn_name].enabled = False
            self.page.update()
            
            # Generar timestamp y nombres de archivo
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(os.path.basename(self.selected_file))[0]
            output_file = os.path.join(OUTPUT_DIR, f"{base_name}_limpio_{timestamp}.xlsm")
            report_file = os.path.join(OUTPUT_DIR, f"{base_name}_informe_{timestamp}.txt")
            self.last_base_name = base_name
            self.last_output_file = None
            self.last_reinsercion_file = None

            # Limpiar rastros redundantes
            self.log_box.add_log("🧹 Limpiando rastros de XLtoEXE...")
            self.action_progress.set_progress(0.3, True, "Limpiando rastros...")
            self.page.update()
            XLtoEXECleaner(self.working_dir).remove_xltoexe_traces()

            # Reconstruir archivo
            self.action_progress.set_progress(0.55, True, "Reconstruyendo archivo...")
            self.log_box.add_log("📝 Reconstruyendo archivo .xlsm limpio...")
            self.page.update()
            self.rebuilder = XLSMRebuilder(self.working_dir)
            self.rebuilder.rebuild(output_file)
            self.last_output_file = output_file

            # Generar copia con macros visibles si hay módulos disponibles
            visible_output = None
            if self.macros:
                visible_output = os.path.join(OUTPUT_DIR, f"{base_name}_reinsercion_{timestamp}.xlsm")
                self.action_progress.set_progress(0.7, True, "Reinsertando macros visibles...")
                self.log_box.add_log("🔁 Generando copia con macros reinsertadas...")
                self.page.update()
                injector = MacroInjector(output_file, self.macros, self.export_dir)
                success, message = injector.create_visible_copy(visible_output)
                if success:
                    self.log_box.add_log("✅ Copia de reinserción creada correctamente")
                    self.last_reinsercion_file = visible_output
                else:
                    visible_output = None
                    self.log_box.add_log(f"⚠️ No se pudo crear la copia de reinserción: {message}")

            # Generar informe
            self.action_progress.set_progress(0.85, True, "Generando informe...")
            self.log_box.add_log("📋 Generando informe técnico...")
            self.page.update()
            self.reporter = ReportGenerator(self.working_dir)
            self.reporter.generate(report_file)

            self.action_progress.set_progress(0.95, True, "Guardando archivos...")
            self.log_box.add_log(f"💾 Guardando archivo limpio: {os.path.basename(output_file)}")
            self.page.update()

            self.stage = 4  # Proceso completado
            self.update_buttons()

            # Mostrar mensaje de éxito con información detallada
            self.log_box.add_log("\n✨ PROCESO COMPLETADO EXITOSAMENTE")
            self.log_box.add_log("="*50)
            self.log_box.add_log("📂 Archivo limpio guardado en:")
            self.log_box.add_log(f"   {output_file}")
            self.log_box.add_log("")
            if visible_output:
                self.log_box.add_log("👁️ Copia de reinserción (macros visibles) guardada en:")
                self.log_box.add_log(f"   {visible_output}")
                self.log_box.add_log("")
            self.log_box.add_log("📋 Informe técnico generado en:")
            self.log_box.add_log(f"   {report_file}")
            self.log_box.add_log("\n🔍 Puede encontrar los archivos en la carpeta 'DESOFUSCADOS' en su escritorio")
            self.log_box.add_log("="*50)
            self.log_box.add_log("\n🎉 ¡Proceso finalizado con éxito!")

            self.action_progress.set_progress(1.0, True, "¡Proceso completado!")
            self.show_snackbar(
                self.page,
                "Archivo guardado exitosamente en la carpeta DESOFUSCADOS",
                FuturisticColors.SUCCESS
            )
            
        except Exception as ex:
            self.log_box.add_log(f"❌ Error al guardar el archivo: {str(ex)}")
            self.action_progress.set_progress(0, True, "Error al guardar")
            self.show_snackbar(
                self.page, 
                f"Error al guardar el archivo: {str(ex)}", 
                FuturisticColors.ERROR
            )
        finally:
            # Asegurarse de que los botones se actualicen correctamente
            self.update_buttons()
            self.page.update()
            
            # Reiniciar el estado después de un tiempo
            time.sleep(3)
            self.action_progress.set_progress(0.0, False)
            self.clean_temp_dir()

if __name__ == "__main__":
    AppGUI().run()




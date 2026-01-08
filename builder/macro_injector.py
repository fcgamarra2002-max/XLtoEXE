import os
import logging
from typing import List, Tuple

try:
    import xlwings as xw
except ImportError:  # pragma: no cover - xlwings is optional at runtime
    xw = None


class MacroInjector:
    """Reconstruye un proyecto VBA visible usando la automatización de Excel."""

    def __init__(self, workbook_path: str, macros: List[dict], export_dir: str | None = None):
        self.workbook_path = workbook_path
        self.macros = macros or []
        self.export_dir = export_dir
        self.last_excel_app = None

    def create_visible_copy(
        self,
        output_path: str,
        *,
        show_excel: bool = False,
        show_vbe: bool = False,
        leave_open: bool = False,
    ) -> Tuple[bool, str]:
        if not self.macros:
            return False, "No hay macros para reinsertar"
        if xw is None:
            return False, "xlwings no está instalado; no se puede automatizar Excel"
        if not os.path.exists(self.workbook_path):
            return False, "No se encontró el archivo fuente para la reinserción"

        app = None
        try:
            app = xw.App(visible=show_excel, add_book=False)
            book = app.books.open(self.workbook_path)
            vb_components = book.api.VBProject.VBComponents

            # Eliminar módulos estándar, clases y formularios existentes
            removable_types = {1, 2, 3}
            to_remove = []
            for i in range(1, vb_components.Count + 1):
                component = vb_components.Item(i)
                if component.Type in removable_types:
                    to_remove.append(component.Name)
            for name in to_remove:
                try:
                    vb_components.Remove(vb_components.Item(name))
                except Exception as exc:  # pragma: no cover - depende de Excel
                    logging.warning("No se pudo eliminar el módulo %s: %s", name, exc)

            # Reinsertar cada módulo según su tipo
            for macro in self.macros:
                module_type = macro.get('type', 'std')
                module_name = macro.get('module_name') or macro.get('filename') or 'Module'
                code = macro.get('code') or ''
                export_path = macro.get('export_path')

                if module_type in ('std', 'class'):
                    self._import_module(vb_components, module_type, module_name, code, export_path)
                elif module_type == 'document':
                    self._replace_document_module(vb_components, module_name, code)
                elif module_type == 'form':
                    logging.warning("Reinserción automática de formularios (%s) no soportada actualmente", module_name)
                else:
                    logging.warning("Tipo de módulo desconocido %s para %s", module_type, module_name)

            # Guardar copia visible
            book.api.SaveCopyAs(output_path)

            if show_vbe:
                try:
                    book.app.api.VBE.MainWindow.Visible = True
                except Exception:  # pragma: no cover - depende de seguridad de Excel
                    logging.warning("No se pudo abrir la ventana del editor VBA automáticamente")

            if leave_open:
                try:
                    preview = book.app.books.open(output_path)
                    preview.activate()
                except Exception:
                    logging.warning("No se pudo abrir la copia de reinserción para revisión manual")
                finally:
                    try:
                        book.close(save=False)
                    except Exception:
                        pass
                self.last_excel_app = app
                return True, "Macros reinsertadas; Excel permanece abierto para revisión"
            else:
                try:
                    book.close(save=False)
                except Exception:
                    pass
            return True, "Macros reinsertadas correctamente"
        except Exception as exc:  # pragma: no cover - depende de Excel
            logging.exception("Fallo al reinsertar macros: %s", exc)
            return False, str(exc)
        finally:
            if app is not None:
                if leave_open:
                    # Dejar Excel abierto para que el usuario revise manualmente
                    self.last_excel_app = app
                else:
                    try:
                        for book in app.books:
                            book.close(save=False)
                    except Exception:
                        pass
                    app.quit()

    @staticmethod
    def _import_module(vb_components, module_type: str, module_name: str, code: str, export_path: str | None):
        # Tipos VBA: 1=Estándar, 2=Clase
        component_type = 1 if module_type == 'std' else 2
        if export_path and os.path.exists(export_path):
            component = vb_components.Import(export_path)
            try:
                component.Name = module_name
            except Exception:
                # Si no se puede renombrar, dejar el nombre generado por Excel
                pass
        else:
            component = vb_components.Add(component_type)
            component.Name = module_name
            code_module = component.CodeModule
            if code_module.CountOfLines:
                code_module.DeleteLines(1, code_module.CountOfLines)
            code_module.AddFromString(code)

    @staticmethod
    def _replace_document_module(vb_components, module_name: str, code: str):
        component = MacroInjector._find_component_by_name(vb_components, module_name)
        if component is None:
            logging.warning("No se encontró el módulo de documento %s para reemplazar", module_name)
            return
        code_module = component.CodeModule
        if code_module.CountOfLines:
            code_module.DeleteLines(1, code_module.CountOfLines)
        code_module.AddFromString(code)

    @staticmethod
    def _find_component_by_name(vb_components, module_name: str):
        for i in range(1, vb_components.Count + 1):
            component = vb_components.Item(i)
            if component.Name.lower() == module_name.lower():
                return component
        return None

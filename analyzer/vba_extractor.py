import os
import pathlib
from oletools.olevba import VBA_Parser

class VBAExtractor:
    def __init__(self, working_dir):
        self.working_dir = working_dir
        self.vba_path = self._find_vba_project()
        self.macros = []

    def _find_vba_project(self):
        # Busca vbaProject.bin en la estructura extraída
        for root, dirs, files in os.walk(self.working_dir):
            for file in files:
                if file.lower() == 'vbaproject.bin':
                    return os.path.join(root, file)
        return None

    def extract_macros(self, export_dir=None):
        if not self.vba_path:
            print('No se encontró vbaProject.bin')
            return []
        vba_parser = VBA_Parser(self.vba_path)
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
            module_name = self._infer_module_name(vba_filename, stream_path)
            module_type = self._infer_module_type(vba_filename, stream_path)
            macro = {
                'filename': vba_filename,
                'stream_path': stream_path,
                'module_name': module_name,
                'type': module_type,
                'code': vba_code,
                'export_path': None,
            }
            if export_dir:
                macro['export_path'] = self._export_macro(macro, export_dir)
            self.macros.append(macro)
        return self.macros

    @staticmethod
    def _export_macro(macro, export_dir):
        os.makedirs(export_dir, exist_ok=True)
        name = macro['module_name'] or macro['filename'] or 'Module'
        safe_name = ''.join(ch if ch.isalnum() or ch in ('_', '-') else '_' for ch in name)
        module_type = macro.get('type', 'std')

        if module_type == 'form':
            extension = '.frm'
        elif module_type in ('class', 'document'):
            extension = '.cls'
        else:
            extension = '.bas'

        filename = safe_name
        if not filename.lower().endswith(extension):
            filename += extension

        path = pathlib.Path(export_dir) / filename
        with open(path, 'w', encoding='utf-8', newline='\r\n') as f:
            f.write(macro['code'])
        return str(path)

    @staticmethod
    def _infer_module_name(vba_filename, stream_path):
        name = vba_filename or ''
        if name.lower().endswith(('.bas', '.cls', '.frm')):
            name = name.rsplit('.', 1)[0]
        if not name:
            name = (stream_path or '').split('/')[-1]
        return name or 'Module'

    @staticmethod
    def _infer_module_type(vba_filename, stream_path):
        name_lower = (vba_filename or '').lower()
        stream_lower = (stream_path or '').lower()
        if name_lower.endswith('.frm') or 'userform' in name_lower:
            return 'form'
        if name_lower.endswith('.cls'):
            return 'class'
        if any(keyword in (name_lower, stream_lower) for keyword in ['thisworkbook', 'sheet', 'hoja']):
            return 'document'
        if name_lower.endswith('.bas'):
            return 'std'
        return 'std'

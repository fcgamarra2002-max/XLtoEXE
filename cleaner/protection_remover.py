import os
import xml.etree.ElementTree as ET

class ProtectionRemover:
    def __init__(self, working_dir):
        self.working_dir = working_dir

    def remove_sheet_and_workbook_protection(self):
        # Elimina protección de workbook y hojas
        for root, dirs, files in os.walk(self.working_dir):
            for file in files:
                if file.endswith('.xml'):
                    path = os.path.join(root, file)
                    self._clean_xml(path)

    def _clean_xml(self, path):
        ET.register_namespace('', "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        tree = ET.parse(path)
        root = tree.getroot()
        changed = False
        for tag in ['{http://schemas.openxmlformats.org/spreadsheetml/2006/main}workbookProtection',
                    '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetProtection']:
            for elem in root.findall(tag):
                root.remove(elem)
                changed = True
        if changed:
            tree.write(path, encoding='utf-8', xml_declaration=True)

    def remove_vba_project_password(self):
        """Parches suaves sobre PROJECT stream para deshabilitar la contraseña sin corromper el binario."""
        vba_path = self._find_vba_project_path()
        if not vba_path:
            return False

        with open(vba_path, 'rb') as f:
            data = f.read()

        replacements = [
            (b'DPB="', b'DPX="'),
            (b'DPc="', b'DPx="'),
            (b'DPB=', b'DPX='),
            (b'DPc=', b'DPx='),
            (b'DPC="', b'DPX="'),
            (b'DPC=', b'DPX='),
            (b'DPb="', b'DPx="'),
            (b'DPb=', b'DPx='),
            (b'CMG="', b'CMX="'),
            (b'CMG=', b'CMX='),
            (b'CMg="', b'CMx="'),
            (b'CMg=', b'CMx='),
            (b'OEMPassword="', b'OEMPassword=""'),
            (b'Password="', b'Password=""'),
        ]

        changed = False
        for old, new in replacements:
            if old in data:
                data = data.replace(old, new)
                changed = True

        if not changed:
            return False

        with open(vba_path, 'wb') as f:
            f.write(data)
        return True

    def _find_vba_project_path(self):
        for root, dirs, files in os.walk(self.working_dir):
            for file in files:
                if file.lower() == 'vbaproject.bin':
                    return os.path.join(root, file)
        return None

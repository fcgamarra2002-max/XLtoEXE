import os
import re

class XLtoEXECleaner:
    def __init__(self, working_dir):
        self.working_dir = working_dir

    def remove_xltoexe_traces(self):
        # Elimina módulos, propiedades y rastros de XLtoEXE
        # 1. Limpiar propiedades customizadas (custom.xml)
        custom_xml = os.path.join(self.working_dir, 'docProps', 'custom.xml')
        if os.path.exists(custom_xml):
            with open(custom_xml, 'r', encoding='utf-8') as f:
                xml_content = f.read()
            xml_content = re.sub(r'XLtoEXE', 'XLT_CLEAN', xml_content)
            with open(custom_xml, 'w', encoding='utf-8') as f:
                f.write(xml_content)

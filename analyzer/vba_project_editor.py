import os
import shutil
import tempfile
from oletools.olevba import VBA_Parser

class VBAProjectEditor:
    """
    Permite reemplazar módulos VBA dentro de vbaProject.bin usando un enfoque de extracción y reinserción.
    """
    def __init__(self, working_dir):
        self.working_dir = working_dir
        self.vba_path = self._find_vba_project()

    def _find_vba_project(self):
        for root, dirs, files in os.walk(self.working_dir):
            for file in files:
                if file.lower() == 'vbaproject.bin':
                    return os.path.join(root, file)
        return None

    def extract_modules(self):
        """
        Extrae los módulos VBA como texto.
        """
        if not self.vba_path:
            return []
        vba_parser = VBA_Parser(self.vba_path)
        modules = []
        for (_, _, vba_filename, vba_code) in vba_parser.extract_macros():
            modules.append({'filename': vba_filename, 'code': vba_code})
        return modules

    def replace_modules(self, new_modules):
        """
        Reemplaza el código de los módulos VBA en vbaProject.bin.
        NOTA: Esto requiere reempaquetar el binario OLE. Aquí se implementa una solución básica usando oletools y ms-office-builder.
        """
        # 1. Extraer todo el vbaProject.bin a una carpeta temporal
        import olefile
        import oletools.olevba
        import io
        if not self.vba_path:
            print('No se encontró vbaProject.bin')
            return False
        temp_dir = tempfile.mkdtemp()
        ole = olefile.OleFileIO(self.vba_path)
        for stream in ole.listdir():
            stream_name = '/'.join(stream)
            data = ole.openstream(stream).read()
            out_path = os.path.join(temp_dir, stream_name.replace('/', '_'))
            with open(out_path, 'wb') as f:
                f.write(data)
        ole.close()
        # 2. Reemplazar el texto de los módulos
        for mod in new_modules:
            mod_file = os.path.join(temp_dir, mod['filename'])
            if os.path.exists(mod_file):
                with open(mod_file, 'wb') as f:
                    f.write(mod['code'].encode('utf-8', errors='replace'))
        # 3. Reempaquetar el OLE (requiere ms-office-builder o similar, o hacerlo manualmente)
        # Aquí solo se deja preparado el reemplazo. Para una solución robusta se recomienda usar vbaextract/vbainject o ms-office-builder.
        print('Los módulos han sido reemplazados en la carpeta temporal:', temp_dir)
        print('Para reempaquetar el vbaProject.bin, utilice una herramienta como vbainject, ms-office-builder o similar.')
        return True

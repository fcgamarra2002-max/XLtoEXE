import os
import pefile
import re
import shutil
import tempfile
import zipfile

class EXEDetector:
    def __init__(self, exe_path):
        self.exe_path = exe_path
        self.xlsm_path = None

    def detect_and_extract(self, exe_path=None, working_dir=None):
        """Intenta localizar o extraer un XLSM a partir de un ejecutable."""
        if exe_path:
            self.exe_path = exe_path

        result = {
            "xlsm_path": None,
            "extracted": False,
        }

        if not self.exe_path or not os.path.exists(self.exe_path):
            return result

        # 1) Si el EXE es en realidad un ZIP autoextraíble, descomprimir en working_dir
        if working_dir and self.unzip_if_needed(working_dir):
            candidate = self._find_xlsm_in_dir(working_dir)
            if candidate:
                result["xlsm_path"] = candidate
                result["extracted"] = True
                return result

        # 2) Buscar rutas incrustadas dentro del binario
        embedded_path = self.extract_embedded_xlsm()
        if embedded_path and os.path.exists(embedded_path):
            result["xlsm_path"] = embedded_path
            return result

        # 3) Como último recurso, revisar %TEMP%
        temp_path = self.extract_temp_xlsm()
        if temp_path and os.path.exists(temp_path):
            result["xlsm_path"] = temp_path
            return result

        return result

    def extract_embedded_xlsm(self):
        # Busca patrones de nombres de archivos Excel en el binario
        with open(self.exe_path, 'rb') as f:
            data = f.read()
        xlsm_candidates = re.findall(b'[\w\/:.-]+\.xlsm', data)
        for candidate in xlsm_candidates:
            candidate_str = candidate.decode(errors='ignore')
            if os.path.exists(candidate_str):
                self.xlsm_path = candidate_str
                return self.xlsm_path
        # Alternativamente, usar pyinstxtractor o monitorear %temp%
        return None

    def extract_temp_xlsm(self):
        # Monitorea la carpeta %temp% en busca de archivos xlsm creados por el EXE
        temp_dir = tempfile.gettempdir()
        for fname in os.listdir(temp_dir):
            if fname.endswith('.xlsm'):
                self.xlsm_path = os.path.join(temp_dir, fname)
                return self.xlsm_path
        return None

    def unzip_if_needed(self, output_dir):
        if zipfile.is_zipfile(self.exe_path):
            with zipfile.ZipFile(self.exe_path, 'r') as zip_ref:
                zip_ref.extractall(output_dir)
            return True
        return False

    @staticmethod
    def _find_xlsm_in_dir(directory):
        if not directory or not os.path.isdir(directory):
            return None
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith('.xlsm'):
                    return os.path.join(root, file)
        return None

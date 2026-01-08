import zipfile
import os
from extractor.zip_handler import ZipHandler

class XLSMUnpacker:
    def __init__(self, input_path, output_dir):
        self.input_path = input_path
        self.output_dir = output_dir

    def unpack(self):
        if zipfile.is_zipfile(self.input_path):
            ZipHandler.extract_zip(self.input_path, self.output_dir)
        elif os.path.isdir(self.input_path):
            # Ya está extraído
            pass
        else:
            raise ValueError("El archivo de entrada no es un .xlsm válido ni una carpeta ZIP extraída.")

        # Aquí se pueden agregar rutinas para eliminar rastros de XLtoEXE
        # y preparar la estructura para los siguientes módulos.

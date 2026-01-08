import zipfile
import os

class ZipHandler:
    @staticmethod
    def extract_zip(zip_path, output_dir):
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)

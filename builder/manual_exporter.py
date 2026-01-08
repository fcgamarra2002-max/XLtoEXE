import os
import shutil

class ManualExporter:
    def __init__(self, working_dir):
        self.working_dir = working_dir

    def export(self):
        export_dir = os.path.join(self.working_dir, 'componentes_extraidos')
        os.makedirs(export_dir, exist_ok=True)
        for root, dirs, files in os.walk(self.working_dir):
            for file in files:
                file_path = os.path.join(root, file)
                if not file_path.startswith(export_dir):
                    rel_path = os.path.relpath(file_path, self.working_dir)
                    dest_path = os.path.join(export_dir, rel_path)
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                    shutil.copy2(file_path, dest_path)

import os

class StructureChecker:
    def __init__(self, working_dir):
        self.working_dir = working_dir

    def has_xl_folder(self):
        return os.path.isdir(os.path.join(self.working_dir, 'xl'))

    def has_vba_project(self):
        for root, dirs, files in os.walk(self.working_dir):
            for file in files:
                if file.lower() == 'vbaproject.bin':
                    return True
        return False

    def check_integrity(self):
        # Valida que existan los componentes clave de un xlsm
        xl_dir = os.path.join(self.working_dir, 'xl')
        required = ['workbook.xml', 'sharedStrings.xml']
        for req in required:
            if not os.path.exists(os.path.join(xl_dir, req)):
                return False
        return True

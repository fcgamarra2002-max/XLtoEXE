import os

class ReportGenerator:
    def __init__(self, working_dir):
        self.working_dir = working_dir

    def generate(self, output_path=None):
        report_path = output_path or os.path.join(self.working_dir, 'informe.txt')
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write('INFORME DE DESOFUSCACIÓN Y LIMPIEZA\n')
            f.write('Directorio de trabajo: %s\n' % self.working_dir)
            # Aquí se pueden agregar detalles de los cambios realizados

import zipfile
import os

class XLSMRebuilder:
    def __init__(self, working_dir):
        self.working_dir = working_dir

    def rebuild(self, output_path=None):
        """
        Reconstruye el archivo XLSM a partir de los archivos extraídos.
        
        Args:
            output_path (str, optional): Ruta completa donde guardar el archivo reconstruido.
                                       Si no se especifica, se guarda en el directorio de trabajo.
        """
        if output_path is None:
            output_path = os.path.join(self.working_dir, 'reconstruido.xlsm')
            
        # Asegurarse de que el directorio de salida existe
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        try:
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Primero agregar [Content_Types].xml si existe
                content_types_path = os.path.join(self.working_dir, '[Content_Types].xml')
                if os.path.exists(content_types_path):
                    zipf.write(content_types_path, '[Content_Types].xml')
                
                # Luego agregar _rels/.rels si existe
                rels_dir = os.path.join(self.working_dir, '_rels')
                if os.path.exists(rels_dir):
                    for file in os.listdir(rels_dir):
                        file_path = os.path.join(rels_dir, file)
                        if os.path.isfile(file_path):
                            zipf.write(file_path, f'_rels/{file}')
                
                # Finalmente agregar el resto de los archivos en orden
                for root, dirs, files in os.walk(self.working_dir):
                    for file in sorted(files):  # Ordenar archivos para consistencia
                        file_path = os.path.join(root, file)
                        # Excluir el archivo de salida si está en el directorio
                        if os.path.abspath(file_path) != os.path.abspath(output_path):
                            arcname = os.path.relpath(file_path, self.working_dir)
                            # No agregar archivos ya agregados
                            if arcname not in ['[Content_Types].xml'] and not arcname.startswith('_rels/'):
                                zipf.write(file_path, arcname)
        
        except Exception as e:
            # Si algo sale mal, eliminar el archivo incompleto
            if os.path.exists(output_path):
                os.remove(output_path)
            raise Exception(f"Error al reconstruir el archivo: {str(e)}")

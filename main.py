
import argparse
import os
import sys
import logging
from extractor.xlsm_unpacker import XLSMUnpacker
from builder.xlsm_rebuilder import XLSMRebuilder
from builder.manual_exporter import ManualExporter
from report.report_generator import ReportGenerator
from analyzer.vba_extractor import VBAExtractor
from cleaner.protection_remover import ProtectionRemover
from cleaner.xlt_exe_cleaner import XLtoEXECleaner
from deobfuscator.vba_deobfuscator import VBADeobfuscator
from deobfuscator.vba_optimizer import VBAOptimizer

def setup_logging(output_dir):
    log_path = os.path.join(output_dir, "proceso.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout)
        ]
    )

def extraer_archivo(input_path, output_dir):
    logging.info("Iniciando extracción del archivo.")
    unpacker = XLSMUnpacker(input_path, output_dir)
    unpacker.unpack()

def limpiar_protecciones(output_dir):
    logging.info("Eliminando protecciones y rastros de XLtoEXE.")
    cleaner = ProtectionRemover(output_dir)
    cleaner.remove_sheet_and_workbook_protection()
    cleaner.remove_vba_project_password()
    xlt_cleaner = XLtoEXECleaner(output_dir)
    xlt_cleaner.remove_xltoexe_traces()

def procesar_macros(output_dir):
    logging.info("Extrayendo y desofuscando macros VBA.")
    vba_analyzer = VBAExtractor(output_dir)
    macros = vba_analyzer.extract_macros()
    if macros:
        deobfuscator = VBADeobfuscator(macros)
        deobfuscated_macros = deobfuscator.deobfuscate()
        optimizer = VBAOptimizer(deobfuscated_macros)
        optimized_macros = optimizer.optimize()
        # TODO: Sobrescribir vbaProject.bin con macros optimizados
    else:
        logging.warning("No se encontraron macros VBA para procesar.")

def reconstruir_o_exportar(output_dir, manual):
    if manual:
        logging.info("Extracción manual seleccionada.")
        exporter = ManualExporter(output_dir)
        exporter.export()
    else:
        logging.info("Reconstruyendo archivo .xlsm limpio.")
        rebuilder = XLSMRebuilder(output_dir)
        rebuilder.rebuild()

def generar_informe(output_dir):
    logging.info("Generando informe final.")
    reporter = ReportGenerator(output_dir)
    reporter.generate()

def main():
    parser = argparse.ArgumentParser(description="Desofuscador y extractor de archivos .xlsm protegidos por XLtoEXE.")
    parser.add_argument('input', help='Archivo .xlsm o carpeta ZIP extraída')
    parser.add_argument('-o', '--output', help='Directorio de salida', default='output')
    parser.add_argument('--manual', action='store_true', help='Extraer componentes manualmente en vez de reconstruir el .xlsm')
    args = parser.parse_args()

    os.makedirs(args.output, exist_ok=True)
    setup_logging(args.output)

    try:
        extraer_archivo(args.input, args.output)
        limpiar_protecciones(args.output)
        procesar_macros(args.output)
        reconstruir_o_exportar(args.output, args.manual)
        generar_informe(args.output)
        logging.info("Proceso completado correctamente.")
    except Exception as e:
        logging.exception(f"Error durante el proceso: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()

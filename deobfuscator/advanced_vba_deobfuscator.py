import re

class AdvancedVBADeobfuscator:
    """
    Desofusca nombres de variables, funciones y módulos ofuscados (ej: kqclqcmqcnqcoqcpqcqqcrqcsqctqc) y los reemplaza por nombres legibles en español.
    Omite nombres ya entendibles y evita la letra Ñ.
    """
    def __init__(self, macros):
        self.macros = macros
        self.renaming_map = {}
        self.used_names = set()

    def is_obfuscated(self, name):
        # Considera ofuscado si es una cadena larga, sin vocales, y no es palabra reconocible
        return (len(name) > 8 and not re.search(r'[aeiouáéíóú]', name, re.IGNORECASE) and name.islower())

    def next_var_name(self, prefix):
        idx = 1
        while True:
            name = f"{prefix}{idx}"
            if name not in self.used_names:
                self.used_names.add(name)
                return name
            idx += 1

    def deobfuscate(self):
        deobfuscated_macros = []
        for macro in self.macros:
            code = macro['code']
            code, renaming_map = self._rename_obfuscated_names(code)
            deobfuscated_macros.append({'filename': macro['filename'], 'code': code})
        return deobfuscated_macros

    def _rename_obfuscated_names(self, code):
        # Encuentra nombres tipo kqclqcmqcnqcoqcpqcqqcrqcsqctqc
        pattern = re.compile(r'\b([a-z]{8,})\b', re.IGNORECASE)
        found = set(pattern.findall(code))
        renaming_map = {}
        for name in found:
            if self.is_obfuscated(name) and 'ñ' not in name:
                if name.startswith('mod'):  # módulo
                    new_name = self.next_var_name('modulo')
                elif name.startswith('sub') or name.startswith('fun'):
                    new_name = self.next_var_name('funcion')
                else:
                    new_name = self.next_var_name('variable')
                renaming_map[name] = new_name
                code = re.sub(rf'\b{name}\b', new_name, code)
        return code, renaming_map

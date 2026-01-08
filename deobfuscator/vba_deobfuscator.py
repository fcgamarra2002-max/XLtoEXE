import re
from oletools import mraptor

class VBADeobfuscator:
    def __init__(self, macros):
        self.macros = macros
        self.renaming_map = {}

    def deobfuscate(self):
        deobfuscated_macros = []
        for macro in self.macros:
            code = macro['code']
            code, renaming_map = self._rename_obfuscated_names(code)
            code = self._add_comments(code)
            deobfuscated_macros.append({'filename': macro['filename'], 'code': code})
        return deobfuscated_macros

    def _rename_obfuscated_names(self, code):
        # Busca nombres tipo a1, b2, c3, etc. y los reemplaza por nombres más descriptivos
        pattern = re.compile(r'\b([a-z]{1,2}\d{1,3})\b', re.IGNORECASE)
        found = set(pattern.findall(code))
        renaming_map = {}
        for idx, name in enumerate(found):
            new_name = f'var_{idx+1}'
            renaming_map[name] = new_name
            code = re.sub(rf'\b{name}\b', new_name, code)
        return code, renaming_map

    def _add_comments(self, code):
        # Añade comentarios en líneas sospechosas (muy cortas, llamadas a funciones, etc.)
        commented = []
        for line in code.splitlines():
            if re.match(r'^\s*(On Error|GoTo|Call|Shell|CreateObject)', line, re.IGNORECASE):
                commented.append(f"'{line} ' [Automático: línea potencialmente relevante]")
            else:
                commented.append(line)
        return '\n'.join(commented)

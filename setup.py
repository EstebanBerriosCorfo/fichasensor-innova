from cx_Freeze import setup, Executable
import os
import sys

# Incluye recursos y librerías necesarias
include_files = [
    ('src/assets', 'src/assets'),
    ('src/templates', 'src/templates')
]

# Añade DLL específica si es necesario
dll_path = os.path.join(os.path.dirname(sys.executable), 'Lib', 'site-packages', 'pandas', '_libs', 'window', 'aggregations.cp39-win_amd64.pyd')
if os.path.exists(dll_path):
    include_files.append((dll_path, 'pandas/_libs/window'))

build_options = {
    'packages': ['pandas', 'customtkinter'],
    'includes': [],
    'include_files': include_files,
    'excludes': []
}

base = 'Win32GUI' if sys.platform == 'win32' else None

setup(
    name='GestorFichasSensor',
    version='1.0',
    description='Aplicación de generación de fichas y cartas',
    options={'build_exe': build_options},
    executables=[Executable('FichaSensor.py', base=base, icon='src/assets/favicon.ico')]
)
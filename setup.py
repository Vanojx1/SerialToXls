import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
buildOptions = dict(packages=['xlsxwriter', 'serial'], excludes=[])
msiOptions = dict(initial_target_dir='[ProgramFilesFolder]\\SerialToXls\\', upgrade_code='{89F23835-9AFD-4381-806F-DCB5905D9A63}')

base = 'Win32GUI' if sys.platform == 'win32' else None

executables = [
    Executable('app.py', base=base, targetName='SerialToXls', shortcutName='SerialToXls', shortcutDir='DesktopFolder',)
]

setup(name='SerialToXls',
      version='1.1',
      description='',
      options=dict(build_exe=buildOptions, bdist_msi=msiOptions),
      executables=executables)

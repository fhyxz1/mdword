import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# 项目主文件
main_script = 'main.py'

# 添加字体文件
fonts_datas = [(os.path.join('fonts', 'SourceHanSansSC-Regular-2.otf'), 'fonts')]

hidden_imports = collect_submodules('keyboard') + collect_submodules('pystray')

a = Analysis(
    [main_script],
    pathex=[],
    binaries=[],
    datas=collect_data_files('kivy') + collect_data_files('pystray') + fonts_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='mdword',
    debug=False,
    strip=False,
    upx=True,
    console=False,
)

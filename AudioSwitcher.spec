# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

datas_comtypes, binaries_comtypes, hiddenimports_comtypes = collect_all('comtypes')
datas_pycaw, binaries_pycaw, hiddenimports_pycaw = collect_all('pycaw')
datas_ctk, binaries_ctk, hiddenimports_ctk = collect_all('customtkinter')

a = Analysis(
    ['switch_audio.py'],
    pathex=[],
    binaries=binaries_comtypes + binaries_pycaw + binaries_ctk,
    # app_icon.icoは設定画面のタイトルバー用にexe内にも同梱する
    datas=datas_comtypes + datas_pycaw + datas_ctk + [('app_icon.ico', '.')],
    hiddenimports=[
        'pycaw',
        'pycaw.pycaw',
        'pycaw.constants',
        'pycaw.utils',
    ] + hiddenimports_comtypes + hiddenimports_pycaw + hiddenimports_ctk,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='AudioSwitcher',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPXで圧縮したexeはWindows Defenderなどが「圧縮済みの不審なexe」として
    # 誤検知しやすいため無効にする(サイズは増えるがインストーラーで圧縮される)。
    upx=False,
    # 以下は旧設定(upx=True時代)の名残。upx=Falseなら実質無効
    # PyInstaller + UPX + Tkinterの組み合わせはtcl/tkのDLLが破損・不安定化する
    # 既知の問題があるため（実機でtcl86t.dll内での不定期クラッシュを確認済み）、
    # Tcl/Tk関連DLLとpythonランタイム本体はUPX圧縮の対象から除外する。
    upx_exclude=['tcl86t.dll', 'tk86t.dll', '_tkinter*.pyd', 'python3*.dll', 'zlib1.dll'],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app_icon.ico'],
)

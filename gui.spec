# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['gui.py'],
             pathex=['C:\\Users\\timot\\Documents\\Programming\\python\\WOIDS'],
             binaries=[],
             datas=[('resources/*', 'resources'), ('WOIDS.ui','.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=True)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [('v', None, 'OPTION')],
          name='WOIDS-Sound',
          debug=True,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False)

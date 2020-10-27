# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['main_term.py', 'afd_config.py', 'afd_docxtpl_for_pd.py', 'afd_pd.py', 'afd_ggl.py'],
             pathex=['E:\\python\\AFDsoftAOSR'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='AOSR',
          debug=True,
          bootloader_ignore_signals=True,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )

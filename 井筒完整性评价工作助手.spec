# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['井筒完整性评价工作助手.py'],
             pathex=['H:\\源代码工区@20201012\\#井筒完整性评价工作助手'],
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
          [],
          exclude_binaries=True,
          name='井筒完整性评价工作助手',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False , icon='H:\\源代码工区@20201012\\#井筒完整性评价工作助手\\resources\\ico\\petro.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='井筒完整性评价工作助手')

# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

added_files = [
		 ('work_setting\\work_help.html', 'work_setting'),
		 ('work_setting\\work_setting.txt', 'work_setting'),
		 ('work_setting\\working_hour.log', 'work_setting'),
		 ( 'icon', 'icon' )
         ]

a = Analysis(['working_hours.pyw'],
             pathex=['C:\\py_virtual\\project\\working-hours'],
             binaries=[],
             datas=added_files,
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
          name='working_hours',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False, icon='icon\\Bill.ico' )

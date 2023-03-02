# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['Mailchimp_Updater.py'],
             pathex=['C:\\Users\\Justin Baron\\PycharmProjects\\InfoodlAPI','C:\\Users\\Justin Baron\\PycharmProjects\\InfoodlAPI\\venv\\Lib\\site-packages'],
             binaries=[],
             datas=[],
             hiddenimports=["selenium","tkinter", "pymsgbox"],
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
          a.scripts + [('v', '', 'OPTION')],
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='MailChimp_Updater',
          debug=True,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True)

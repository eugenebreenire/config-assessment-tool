# -*- mode: python ; coding: utf-8 -*-

import os, sys
from os import path
site_packages = next(p for p in sys.path if 'site-packages' in p) # for pptx

block_cipher = None
bundle_name = "config-assessment-tool"
exec_file_name = bundle_name
platform = ""
platform_binaries=[]

if sys.platform == "win32":
    platform = "-windows"
    exec_file_name = f"{bundle_name}.exe"
elif sys.platform == "linux":
    platform = "-linux"
    platform_binaries=[('/usr/local/lib/libcrypt.so.2','.')]
elif sys.platform == "darwin":
    platform = "-macosx"
else:
    print(f"Platform not clear. Creating generic bundle {sys.platform}")

version = open("VERSION", "r").read().strip()
bundle_name = f"{bundle_name}{platform}-{version}"

a = Analysis(
    ["../backend/backend.py"],
    pathex=["./backend", "."],
    binaries=platform_binaries,
    datas=[
        ("../backend/resources/img/splash.txt", "backend/resources/img"),
        ("../VERSION", "."),
        ("../backend/resources/pptAssets/background.jpg", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/background_2.jpg", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/criteria.png", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/criteria2.png", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/checkmark.png", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/xmark.png", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/HybridApplicationMonitoringUseCase.json", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/HybridApplicationMonitoringUseCase_template.pptx", "backend/resources/pptAssets"),
     	(path.join(site_packages,"pptx","templates"), "pptx/templates"), # for pptx

        # Config files previously handled by manual copy
        ("../input/jobs/DefaultJob.json", "input/jobs"),
        ("../input/thresholds/DefaultThresholds.json", "input/thresholds"),
        ("../backend/resources/controllerDefaults/defaultHealthRulesAPM.json", "backend/resources/controllerDefaults"),
        ("../backend/resources/controllerDefaults/defaultHealthRulesBRUM.json", "backend/resources/controllerDefaults"),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=exec_file_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas, strip=False, upx=True, upx_exclude=[], name=bundle_name)


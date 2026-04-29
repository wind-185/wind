# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ["chat_analyzer.py"],
    pathex=[],
    binaries=[],
    datas=[("bee_icon.ico", ".")],
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name="MofangVoiceProcessor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="MofangVoiceProcessor",
)

app = BUNDLE(
    coll,
    name="MofangVoiceProcessor.app",
    bundle_identifier="com.xiaoduo.voice-analyzer",
    info_plist={
        "CFBundleDisplayName": "魔方原声处理",
        "CFBundleName": "魔方原声处理",
        "NSHighResolutionCapable": "True",
    },
)

# -*- mode: python ; coding: utf-8 -*-
"""
PoSmelter_License_Tool.spec
PyInstaller 빌드 설정
"""
import sys
import os
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

# ──────────────────────────────────────────────
# 패키지 자동 수집
# ──────────────────────────────────────────────
datas     = []
binaries  = []
hiddenimports = []

for pkg in ("streamlit", "streamlit.web", "streamlit.runtime",
            "altair", "pandas", "numpy", "reportlab", "pptx"):
    d, b, h = collect_all(pkg)
    datas    += d
    binaries += b
    hiddenimports += h

# 앱 파일 포함 (app.py는 _MEIPASS 루트에 배치)
datas += [
    ("app.py", "."),
    ("components", "components"),
    ("라이선스_확인서_템플릿.pptx", "."),
]

# 한글 폰트 포함 (Windows Malgun Gothic)
malgun = r"C:\Windows\Fonts\malgun.ttf"
malgunbd = r"C:\Windows\Fonts\malgunbd.ttf"
if os.path.exists(malgun):
    datas += [(malgun, "fonts")]
if os.path.exists(malgunbd):
    datas += [(malgunbd, "fonts")]

# streamlit static 파일
try:
    import streamlit
    st_root = os.path.dirname(streamlit.__file__)
    datas += [(os.path.join(st_root, "static"),  "streamlit/static")]
    datas += [(os.path.join(st_root, "runtime"), "streamlit/runtime")]
except Exception:
    pass

hiddenimports += [
    "streamlit",
    "streamlit.web.cli",
    "streamlit.web.server",
    "click",
    "pandas",
    "numpy",
    "reportlab",
    "reportlab.pdfbase.ttfonts",
    "reportlab.lib.pagesizes",
    "pptx",
    "python_pptx",
    "altair",
    "packaging",
    "pyarrow",
    "pyarrow.vendored.version",
    "tzdata",
    "importlib_metadata",
    "pkg_resources",
    "tomli",
    "typing_extensions",
    "watchdog",
    "validators",
]

# ──────────────────────────────────────────────
# Analysis
# ──────────────────────────────────────────────
a = Analysis(
    ["launcher.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "scipy", "PIL", "cv2", "tkinter"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="AnsysLicenseTool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # 콘솔 창 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="icon.ico",      # 아이콘 파일이 있으면 주석 해제
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="AnsysLicenseTool",
)

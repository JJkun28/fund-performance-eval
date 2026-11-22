"""用真实 Excel（Windows，pywin32）或 LibreOffice 重算公式，返回带缓存值的副本路径。"""
import os
import shutil
import subprocess
import sys


def _soffice():
    for p in (shutil.which('soffice'), shutil.which('libreoffice'),
              r'C:\Program Files\LibreOffice\program\soffice.exe',
              '/Applications/LibreOffice.app/Contents/MacOS/soffice'):
        if p and os.path.exists(p):
            return p
    return None


def available_engine():
    want = os.environ.get('FPE_RECALC', 'auto')
    if want in ('auto', 'excel') and sys.platform == 'win32':
        try:
            import win32com.client  # noqa: F401
            return 'excel'
        except ImportError:
            if want == 'excel':
                return None
    if want in ('auto', 'libreoffice') and _soffice():
        return 'libreoffice'
    return None


def recalc(path: str, out_dir: str) -> str:
    engine = available_engine()
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, os.path.basename(path))
    if engine == 'excel':
        import win32com.client
        xl = win32com.client.DispatchEx('Excel.Application')
        xl.Visible = False
        xl.DisplayAlerts = False
        try:
            wb = xl.Workbooks.Open(os.path.abspath(path))
            xl.CalculateFull()
            wb.SaveAs(os.path.abspath(out), FileFormat=51)
            wb.Close(False)
        finally:
            xl.Quit()
        return out
    if engine == 'libreoffice':
        subprocess.run([_soffice(), '--headless', '--calc', '--convert-to', 'xlsx',
                        '--outdir', out_dir, path], check=True, capture_output=True, timeout=180)
        return out
    raise RuntimeError('没有可用的重算引擎（Excel 或 LibreOffice）')

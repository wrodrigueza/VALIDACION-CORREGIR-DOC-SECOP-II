# -*- coding: utf-8 -*-
"""
SECOP II DOC CHECK & FIX
- Sanea nombres (minúsculas + dígitos, sin espacios ni símbolos)
- Fusiona subcarpetas por prefijo saneado de 10
- Sufijo 'C' en TODAS las carpetas/archivos (en la base del nombre)
- Convierte no-PDF a PDF; si NO se puede convertir, se EXTRAEN a "FORMATO DIFERENTE A PDF"
- Copia no bloqueante por bloques (con timeout y pregunta si continuar)
- Barra de progreso global + barra de progreso por archivo
- Burbujeo por MaxPath y MaxDepth
- Elimina carpetas vacías (corregida y dump)
- Reportes HTML: INICIAL, CORREGIDO, FINAL (incluye “Archivos omitidos”)
"""

import os, re, sys, platform, webbrowser, subprocess, time, shutil, unicodedata, glob, tempfile, textwrap, importlib.util
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, Callable
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ======= Fix Tkinter en ejecutables PyInstaller (opcional) =======
def _tk_fix_meipass():
    """Ajusta TCL/TK si el exe corre con --onefile. No es obligatorio con _tk_data, pero no estorba."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = os.path.join(sys._MEIPASS, "_tk_data")
        for pat, env in (("tcl8.*", "TCL_LIBRARY"), ("tk8.*", "TK_LIBRARY")):
            hits = sorted(glob.glob(os.path.join(base, pat)), reverse=True)
            if hits:
                os.environ[env] = hits[0]
_tk_fix_meipass()

# ======= Config =======
PROJECT_NAME = "SECOP II DOC CHECK"
REPORT_FOLDER_NAME = "SECOP_DOC_CHECK"

MAX_PATH_DEFAULT = 240
MAX_FILE_NAME_DEFAULT = 40   # Límite solicitado
MAX_DEPTH_DEFAULT = 5

KEEP_SHORT_FILE = 10
KEEP_SHORT_DIR  = 10
SUFFIX_C = "C"

GENERATE_CSV_REPORTS = False
GENERATE_MAPPING_CSV = False

NON_PDF_DUMP_NAME = "FORMATO DIFERENTE A PDF"

# Timeouts conversores
TIMEOUT_SOFFICE = 180
TIMEOUT_CHROME  = 120
TIMEOUT_WKHTML  = 90

# Copia no bloqueante
CHUNK_SIZE   = 1 * 1024 * 1024   # 1 MB -> UI mucho más fluida
TIMEOUT_COPY = 120               # s antes de preguntar por copia lenta

# Validación
FORBIDDEN_CHARS_PATTERN = r'[\\/:*?"<>|%&#+\{\}\[\];,=]'
DIACRITIC_CHARS_PATTERN  = r'[áéíóúüÁÉÍÓÚÜñÑçÇãÃõÕ]'
HIDDEN_BASENAMES = {'Thumbs.db', '.DS_Store', '.ds_store', 'desktop.ini'}

FORBIDDEN_RE = re.compile(FORBIDDEN_CHARS_PATTERN)
DIACRITICS_RE = re.compile(DIACRITIC_CHARS_PATTERN, re.UNICODE)

IMG_EXTS  = {'.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.gif', '.webp'}
TXT_EXTS  = {'.txt', '.csv', '.md', '.log'}
HTML_EXTS = {'.html', '.htm'}
OFFICE_EXTS = {'.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.odt', '.ods', '.odp', '.rtf'}

# ======= UI helpers =======
def _ui_pump():
    """Bombea la UI agresivamente para evitar 'Not Responding'."""
    try:
        if tk._default_root is not None:
            tk._default_root.update_idletasks()
            tk._default_root.update()
    except Exception:
        pass

def run_cmd_with_timeout_ex(cmd, timeout_sec: int) -> Tuple[bool, bool]:
    """Ejecuta comando externo con bombeo de UI. -> (ok, timed_out)."""
    try:
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception:
        return False, False
    t0 = time.time()
    while True:
        _ui_pump()
        ret = p.poll()
        if ret is not None:
            return (ret == 0), False
        if time.time() - t0 > timeout_sec:
            try: p.kill()
            except Exception: pass
            return False, True
        time.sleep(0.05)

# ======= Path helpers =======
def safe_walk(root: str):
    """os.walk tolerante a rutas largas en Windows."""
    if platform.system() == 'Windows':
        def add_prefix(p):
            p = os.path.abspath(p)
            if p.startswith('\\\\?\\') or p.startswith('\\\\.\\'): return p
            if p.startswith('\\\\'): return '\\\\?\\UNC\\' + p[2:]
            return '\\\\?\\' + p
        try:
            for dirpath, dirnames, filenames in os.walk(add_prefix(root)):
                if dirpath.startswith('\\\\?\\UNC\\'):
                    dirpath = '\\\\' + dirpath[8:]
                elif dirpath.startswith('\\\\?\\'):
                    dirpath = dirpath[4:]
                yield dirpath, dirnames, filenames
            return
        except Exception:
            pass
    for dirpath, dirnames, filenames in os.walk(root):
        yield dirpath, dirnames, filenames

def longpath(p: Path) -> str:
    if platform.system() != 'Windows': return str(p)
    ab = os.path.abspath(str(p))
    if ab.startswith('\\\\?\\') or ab.startswith('\\\\.\\'): return ab
    if ab.startswith('\\\\'): return '\\\\?\\UNC\\' + ab[2:]
    return '\\\\?\\' + ab

# ======= Nombre saneado =======
def remove_diacritics(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if not unicodedata.combining(c))

def sanitize_component_letters_digits(name: str) -> str:
    s = remove_diacritics(name).lower()
    s = s.replace(' ', '')
    s = re.sub(r'[^a-z0-9]', '', s)
    return s or 'a'

sanitize_component_strict = sanitize_component_letters_digits

def limit_filename(name: str, max_len: int) -> str:
    if len(name) <= max_len: return name
    base, ext = os.path.splitext(name)
    keep = max_len - len(ext)
    if keep < 1: return (base + ext)[:max_len]
    return (base[:keep] or "a") + ext

def ensure_unique_preserving_C(target_dir: Path, name: str) -> str:
    base, ext = (os.path.splitext(name) if "." in name and not name.startswith(".") else (name, ""))
    has_c = base.endswith(SUFFIX_C)
    core = base[:-len(SUFFIX_C)] if has_c else base
    suffix = SUFFIX_C if has_c else ""
    cand = limit_filename(core + suffix + ext, MAX_FILE_NAME_DEFAULT)
    i = 2
    while (target_dir / cand).exists():
        cand = limit_filename(f"{core}{i}{suffix}{ext}", MAX_FILE_NAME_DEFAULT); i += 1
    return cand

def ensure_unique_generic(target_dir: Path, name: str) -> str:
    base, ext = (os.path.splitext(name) if "." in name and not name.startswith(".") else (name, ""))
    cand = limit_filename(base + ext, MAX_FILE_NAME_DEFAULT)
    i = 2
    while (target_dir / cand).exists():
        cand = limit_filename(f"{base}{i}{ext}", MAX_FILE_NAME_DEFAULT); i += 1
    return cand

# ======= Métricas relativas =======
def _rel_len(full_path: str, root: str) -> int:
    rel = os.path.relpath(full_path, root)
    return 0 if rel in ('.', '') else len(rel)

def _rel_join_len(dir_path: Path, name: str, base_root: Path) -> int:
    rel_dir = os.path.relpath(str(dir_path), str(base_root))
    if rel_dir in ('.', ''):
        return len(name)
    return len(rel_dir) + 1 + len(name)

def rel_depth(full_path: str, root: str) -> int:
    rel = os.path.relpath(full_path, root)
    if rel in ('.', ''): return 0
    return len([p for p in re.split(r'[\\/]+', rel) if p])

def bubble_dir_for_maxdepth(target_dir: Path, floor_dir: Path) -> Path:
    while rel_depth(str(target_dir), str(floor_dir)) > MAX_DEPTH_DEFAULT and target_dir != floor_dir:
        target_dir = target_dir.parent
    return target_dir

def bubble_file_for_maxdepth(target_dir: Path, floor_dir: Path) -> Path:
    while (rel_depth(str(target_dir), str(floor_dir)) + 1) > MAX_DEPTH_DEFAULT and target_dir != floor_dir:
        target_dir = target_dir.parent
    return target_dir

def fit_in_maxpath_bubbling(target_dir: Path, floor_dir: Path, name: str, keep_C_suffix: bool):
    def _ensure_unique(d: Path, n: str) -> str:
        return ensure_unique_preserving_C(d, n) if keep_C_suffix else ensure_unique_generic(d, n)

    cur_dir, cur_name = target_dir, name
    while _rel_join_len(cur_dir, cur_name, floor_dir) > MAX_PATH_DEFAULT and cur_dir != floor_dir:
        cur_dir = cur_dir.parent
        cur_name = _ensure_unique(cur_dir, cur_name)

    if _rel_join_len(cur_dir, cur_name, floor_dir) > MAX_PATH_DEFAULT:
        rel_dir = os.path.relpath(str(cur_dir), str(floor_dir))
        rel_dir_len = 0 if rel_dir in ('.', '') else len(rel_dir)
        sep = 0 if rel_dir_len == 0 else 1
        allowed = MAX_PATH_DEFAULT - rel_dir_len - sep
        if allowed < 1: allowed = 1

        base, ext = (os.path.splitext(cur_name) if "." in cur_name and not cur_name.startswith(".") else (cur_name, ""))
        if keep_C_suffix:
            has_c = base.endswith(SUFFIX_C)
            core = base[:-len(SUFFIX_C)] if has_c else base
            keep = max(1, allowed - len(ext) - len(SUFFIX_C))
            cur_name = (core[:keep] or "a") + SUFFIX_C + ext
        else:
            keep = max(1, allowed - len(ext))
            cur_name = (base[:keep] or "a") + ext
        cur_name = _ensure_unique(cur_dir, cur_name)
    return cur_dir, cur_name

def fit_dirname_in_maxpath_bubbling(parent_dir: Path, floor_dir: Path, name: str):
    return fit_in_maxpath_bubbling(parent_dir, floor_dir, name, keep_C_suffix=True)

# ======= Validación =======
def _validate_item(name, full, tipo, root, counts):
    plen = _rel_len(full, root)
    nlen = len(name)
    depth = rel_depth(full, root)

    issues = []
    if FORBIDDEN_RE.search(name):
        issues.append('CaracteresProhibidos'); counts['forbidden_chars'] += 1
    if DIACRITICS_RE.search(name):
        issues.append('Diacríticos'); counts['diacritics'] += 1
    if plen > MAX_PATH_DEFAULT:
        issues.append('Ruta>MaxPath'); counts['too_long_path'] += 1
    if nlen > MAX_FILE_NAME_DEFAULT:
        issues.append('Nombre>MaxFileName'); counts['too_long_name'] += 1
    if depth > MAX_DEPTH_DEFAULT:
        issues.append('Profundidad>MaxDepth'); counts['too_deep'] += 1
    if name in HIDDEN_BASENAMES or name.startswith('~$'):
        issues.append('Oculto/Temporales'); counts['hidden_temp'] += 1

    order = {'CaracteresProhibidos': 0, 'Diacríticos': 1, 'Ruta>MaxPath': 2,
             'Nombre>MaxFileName': 3, 'Profundidad>MaxDepth': 4, 'Oculto/Temporales': 5}
    issues.sort(key=lambda x: order.get(x, 99))
    return {'Tipo': tipo, 'Ruta': full, 'Nombre': name,
            'LongRuta': plen, 'LongNombre': nlen,
            'Profundidad': depth, 'Problemas': ','.join(issues)}

def validate_folder(source_folder):
    results = []
    counts = {'too_long_path': 0, 'too_long_name': 0, 'too_deep': 0,
              'forbidden_chars': 0, 'diacritics': 0, 'hidden_temp': 0}
    source_folder = os.path.abspath(source_folder)
    for dirpath, dirnames, filenames in safe_walk(source_folder):
        current_dir_name = os.path.basename(dirpath)
        if current_dir_name:
            results.append(_validate_item(current_dir_name, dirpath, 'DIR', source_folder, counts))
        for d in dirnames:
            results.append(_validate_item(d, os.path.join(dirpath, d), 'DIR', source_folder, counts))
        for fn in filenames:
            results.append(_validate_item(fn, os.path.join(dirpath, fn), 'FILE', source_folder, counts))
    def prio(row):
        probs = row.get('Problemas') or ''
        keys = ['CaracteresProhibidos','Diacríticos','Ruta>MaxPath','Nombre>MaxFileName','Profundidad>MaxDepth','Oculto/Temporales']
        idx = min([keys.index(k) for k in keys if k in probs] or [99])
        return (idx, -row.get('LongRuta', 0))
    results.sort(key=prio)
    return results, counts

# ======= Conversión a PDF =======
def which(cmd: str):
    from shutil import which as _which
    return _which(cmd)

def chrome_exe_guess():
    candidates = ["chrome", "google-chrome", "chromium", "chromium-browser", "msedge"]
    for c in candidates:
        p = which(c)
        if p: return p
    win_candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    for p in win_candidates:
        if Path(p).exists(): return p
    return None

def convert_office_to_pdf(src: Path, out_pdf: Path) -> Tuple[bool, bool]:
    soffice = which("soffice") or which("libreoffice")
    if soffice:
        try:
            cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_pdf.parent), str(src)]
            ok, to = run_cmd_with_timeout_ex(cmd, TIMEOUT_SOFFICE)
            if not ok: return False, to
            produced = out_pdf.parent / (src.stem + ".pdf")
            if produced.exists():
                produced.rename(out_pdf); return True, False
        except Exception:
            pass
    if platform.system() == "Windows":
        try:
            import win32com.client as win32  # type: ignore
            ext = src.suffix.lower()
            if ext in {'.doc', '.docx', '.rtf'}:
                word = win32.DispatchEx('Word.Application'); word.Visible = False
                doc = word.Documents.Open(str(src))
                doc.ExportAsFixedFormat(str(out_pdf), 17); doc.Close(False); word.Quit(); return True, False
            if ext in {'.xls', '.xlsx'}:
                excel = win32.DispatchEx('Excel.Application'); excel.Visible = False
                wb = excel.Workbooks.Open(str(src))
                wb.ExportAsFixedFormat(0, str(out_pdf)); wb.Close(False); excel.Quit(); return True, False
            if ext in {'.ppt', '.pptx'}:
                ppt = win32.DispatchEx('PowerPoint.Application'); ppt.Visible = False
                pres = ppt.Presentations.Open(str(src), WithWindow=False)
                pres.SaveAs(str(out_pdf), 32); pres.Close(); ppt.Quit(); return True, False
        except Exception:
            pass
    return False, False

def convert_image_to_pdf(src: Path, out_pdf: Path) -> Tuple[bool, bool]:
    try:
        import img2pdf  # type: ignore
        with open(out_pdf, "wb") as f:
            f.write(img2pdf.convert(str(src)))
        return True, False
    except Exception:
        pass
    try:
        from PIL import Image, ImageSequence  # type: ignore
    except Exception:
        return False, False
    try:
        img = Image.open(src)
        if src.suffix.lower() in {'.tif', '.tiff', '.gif'}:
            frames = [frame.convert("RGB") for frame in ImageSequence.Iterator(img)] or [img.convert("RGB")]
            frames[0].save(out_pdf, save_all=True, append_images=frames[1:])
        else:
            img.convert("RGB").save(out_pdf)
        return True, False
    except Exception:
        return False, False

def convert_text_to_pdf(src: Path, out_pdf: Path) -> Tuple[bool, bool]:
    try:
        from reportlab.lib.pagesizes import A4  # type: ignore
        from reportlab.pdfgen import canvas     # type: ignore
    except Exception:
        return False, False
    try:
        text = src.read_text(encoding='utf-8', errors='ignore').splitlines()
        c = canvas.Canvas(str(out_pdf), pagesize=A4)
        w, h = A4; x, y, dy = 40, h-40, 12
        for line in text:
            c.drawString(x, y, line[:120]); y -= dy
            if y < 40: c.showPage(); y = h-40
        c.save(); return True, False
    except Exception:
        return False, False

def convert_html_to_pdf(src: Path, out_pdf: Path) -> Tuple[bool, bool]:
    chrome = chrome_exe_guess()
    if chrome:
        try:
            cmd = [chrome, "--headless", "--disable-gpu", f"--print-to-pdf={str(out_pdf)}", str(src.resolve().as_uri())]
            ok, to = run_cmd_with_timeout_ex(cmd, TIMEOUT_CHROME)
            if ok and out_pdf.exists(): return True, False
            if to: return False, True
        except Exception:
            pass
    wk = which("wkhtmltopdf")
    if wk:
        try:
            ok, to = run_cmd_with_timeout_ex([wk, str(src), str(out_pdf)], TIMEOUT_WKHTML)
            if ok and out_pdf.exists(): return True, False
            if to: return False, True
        except Exception:
            pass
    return False, False

def convert_any_to_pdf(src: Path, out_pdf: Path) -> Tuple[bool, bool]:
    ext = src.suffix.lower()
    if ext == ".pdf":
        try:
            shutil.copy2(longpath(src), longpath(out_pdf)); return True, False
        except Exception:
            return False, False
    if ext in IMG_EXTS:  return convert_image_to_pdf(src, out_pdf)
    if ext in TXT_EXTS:  return convert_text_to_pdf(src, out_pdf)
    if ext in HTML_EXTS: return convert_html_to_pdf(src, out_pdf)
    if ext in OFFICE_EXTS: return convert_office_to_pdf(src, out_pdf)
    soffice = which("soffice") or which("libreoffice")
    if soffice:
        try:
            cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_pdf.parent), str(src)]
            ok, to = run_cmd_with_timeout_ex(cmd, TIMEOUT_SOFFICE)
            if not ok: return False, to
            produced = out_pdf.parent / (src.stem + ".pdf")
            if produced.exists():
                produced.rename(out_pdf); return True, False
        except Exception:
            pass
    return False, False

# ======= Conteos/tamaños =======
def count_files(root: Path) -> int:
    total = 0
    for dirpath, _, filenames in safe_walk(str(root)):
        for fn in filenames:
            if fn in HIDDEN_BASENAMES or fn.startswith('~$'): continue
            total += 1
    return total

def _file_size_bytes(p: Path) -> int:
    try:
        return p.stat().st_size
    except Exception:
        try:
            return os.stat(longpath(p)).st_size
        except Exception:
            return 0

def dir_size_bytes(root: Path) -> int:
    total = 0
    for dirpath, _, filenames in safe_walk(str(root)):
        for fn in filenames:
            if fn in HIDDEN_BASENAMES or fn.startswith('~$'):
                continue
            total += _file_size_bytes(Path(dirpath) / fn)
    return total

def human_size(n: int) -> str:
    units = ["B","KB","MB","GB","TB","PB"]
    f = float(n)
    for u in units:
        if f < 1024.0:
            return f"{f:.2f} {u}"
        f /= 1024.0
    return f"{f:.2f} EB"

# ======= Eliminar carpetas vacías =======
def prune_empty_dirs(root: Path, keep_root: bool = True) -> int:
    removed = 0
    for dirpath, dirnames, filenames in os.walk(root, topdown=False):
        p = Path(dirpath)
        if keep_root and p == root:
            continue
        try:
            if not any(p.iterdir()):
                p.rmdir()
                removed += 1
        except Exception:
            pass
    return removed

# ======= Copia/corrección =======
def copy_file_chunked(src: Path, dst: Path, timeout_sec: int | None,
                      inner_cb: Optional[Callable[[int, int], None]] = None) -> tuple[bool, bool]:
    """Copia en bloques con UI y callback de progreso por archivo. -> (ok, timed_out)."""
    total = _file_size_bytes(src)
    copied = 0
    start = time.time()
    try:
        with open(longpath(src), 'rb') as fsrc, open(longpath(dst), 'wb') as fdst:
            while True:
                buf = fsrc.read(CHUNK_SIZE)
                if not buf:
                    break
                fdst.write(buf)
                copied += len(buf)
                if inner_cb:
                    try: inner_cb(copied, total)
                    except Exception: pass
                _ui_pump()
                if timeout_sec is not None and (time.time() - start) > timeout_sec:
                    return False, True
        try:
            shutil.copystat(longpath(src), longpath(dst), follow_symlinks=True)
        except Exception:
            pass
        if inner_cb:
            try: inner_cb(total, total)
            except Exception: pass
        return True, False
    except Exception:
        return False, False

def copy_with_prompt_on_timeout(src: Path, dst: Path, relq: str,
                                inner_cb: Optional[Callable[[int,int], None]] = None) -> str:
    ok, to = copy_file_chunked(src, dst, TIMEOUT_COPY, inner_cb=inner_cb)
    if ok:
        return "COPIADO"
    if to:
        ans = messagebox.askyesno(
            PROJECT_NAME,
            f"La copia de este archivo está tardando demasiado:\n\n{relq}\n\n"
            f"¿Quieres CONTINUAR la copia sin límite (ventana fluida) a la carpeta CORREGIDA?\n\n"
            f"Sí = seguir copiando\nNo = omitir"
        )
        if ans:
            ok2, _ = copy_file_chunked(src, dst, None, inner_cb=inner_cb)  # sin límite
            return "COPIADO_LENTO" if ok2 else "ERROR"
        else:
            return "OMITIDO_STUCK"
    return "ERROR"

def copy_with_rules_and_convert(src_root: Path, out_parent: Path, progress_cb=None, total_files: Optional[int] = None,
                                file_progress_cb: Optional[Callable[[int,int,str], None]] = None):
    mapping = []

    sane_root = sanitize_component_letters_digits(src_root.name)
    root_candidate = f"{sane_root}corregido{SUFFIX_C}"
    root_name_out = ensure_unique_preserving_C(out_parent, root_candidate)
    base_dest = out_parent / root_name_out
    base_dest.mkdir(parents=True, exist_ok=True)

    dump_dir = out_parent / NON_PDF_DUMP_NAME
    dump_dir.mkdir(parents=True, exist_ok=True)

    dir_map: dict[Path, Path] = {src_root: base_dest}
    merge_map: dict[tuple[str, str], Path] = {}

    processed = 0
    total = total_files if total_files is not None else count_files(src_root)

    for dirpath, _, filenames in safe_walk(str(src_root)):
        cur = Path(dirpath)

        if cur != src_root:
            src_parent = cur.parent
            dest_parent = dir_map[src_parent]
            dest_parent = bubble_file_for_maxdepth(dest_parent, base_dest)

            clean = sanitize_component_letters_digits(cur.name)
            key10 = (clean[:KEEP_SHORT_DIR] or "a")
            name10C = limit_filename(key10 + SUFFIX_C, MAX_FILE_NAME_DEFAULT)

            candidate = ensure_unique_preserving_C(dest_parent, name10C)
            dest_parent_final, candidate_final = fit_dirname_in_maxpath_bubbling(dest_parent, base_dest, candidate)
            dest_dir = dest_parent_final / candidate_final

            merge_key = (str(dest_parent_final), key10)
            if merge_key in merge_map:
                dest_dir = merge_map[merge_key]
            else:
                dest_dir = bubble_dir_for_maxdepth(dest_dir, base_dest)
                dest_dir.mkdir(parents=True, exist_ok=True)
                merge_map[merge_key] = dest_dir

            dir_map[cur] = dest_dir
        else:
            dest_dir = base_dest

        mapping.append({"Tipo":"DIR","Original":str(cur), "Corregido":str(dir_map[cur]), "Estado":"OK"})

        for fn in filenames:
            if fn in HIDDEN_BASENAMES or fn.startswith('~$'):
                mapping.append({"Tipo":"FILE","Original":str(cur / fn), "Corregido":"--OMITIDO (oculto/temp)", "Estado":"OMITIDO"})
                continue

            src_file = cur / fn
            base, ext = os.path.splitext(fn)
            base_clean = sanitize_component_letters_digits(base)
            ext_lower  = ext.lower()

            # Progreso global
            processed += 1
            if progress_cb:
                try:
                    rel_disp = str(src_file.relative_to(src_root)).replace('\\', '/')
                except Exception:
                    rel_disp = src_file.name
                try: progress_cb(processed, total, rel_disp)
                except Exception: pass

            # Callback para barra por archivo
            def inner_cb(done: int, tot: int):
                if file_progress_cb:
                    file_progress_cb(done, tot, src_file.name)

            if ext_lower == ".pdf":
                out_name = limit_filename(base_clean + SUFFIX_C + ext_lower, MAX_FILE_NAME_DEFAULT)
                out_name = ensure_unique_preserving_C(dir_map[cur], out_name)

                depth_ok_dir = bubble_file_for_maxdepth(dir_map[cur], base_dest)
                dest_dir_final, out_name_final = fit_in_maxpath_bubbling(
                    target_dir=depth_ok_dir, floor_dir=base_dest, name=out_name, keep_C_suffix=True
                )
                target_path = dest_dir_final / out_name_final
                target_path.parent.mkdir(parents=True, exist_ok=True)

                try:
                    try:
                        relq = str(src_file.relative_to(src_root)).replace('\\', '/')
                    except Exception:
                        relq = src_file.name
                    estado = copy_with_prompt_on_timeout(src_file, target_path, relq, inner_cb=inner_cb)
                    if estado.startswith("COPIADO"):
                        mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":str(target_path), "Estado":estado})
                    elif estado == "OMITIDO_STUCK":
                        mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":"--OMITIDO (copia lenta)", "Estado":"OMITIDO_STUCK"})
                    else:
                        mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":"--ERROR_COPIA", "Estado":"ERROR"})
                except Exception as e:
                    mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":f"--ERROR_COPIA: {e}", "Estado":"ERROR"})

            else:
                out_name_pdf = limit_filename(base_clean + SUFFIX_C + ".pdf", MAX_FILE_NAME_DEFAULT)
                out_name_pdf = ensure_unique_preserving_C(dir_map[cur], out_name_pdf)

                depth_ok_dir = bubble_file_for_maxdepth(dir_map[cur], base_dest)
                pdf_dir_final, out_name_pdf_final = fit_in_maxpath_bubbling(
                    target_dir=depth_ok_dir, floor_dir=base_dest, name=out_name_pdf, keep_C_suffix=True
                )
                pdf_target = pdf_dir_final / out_name_pdf_final
                pdf_target.parent.mkdir(parents=True, exist_ok=True)

                converted, timed_out = False, False
                try:
                    converted, timed_out = convert_any_to_pdf(src_file, pdf_target)
                except Exception:
                    converted, timed_out = (False, False)

                if converted:
                    mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":str(pdf_target), "Estado":"CONVERTIDO"})
                else:
                    if timed_out:
                        try:
                            relq = str(src_file.relative_to(src_root)).replace('\\', '/')
                        except Exception:
                            relq = src_file.name
                        ans = messagebox.askyesno(
                            PROJECT_NAME,
                            f"El archivo tardó demasiado o el conversor se quedó pegado:\n\n{relq}\n\n"
                            f"¿Quieres COPIAR el archivo original (sin convertir) a la carpeta CORREGIDA?\n\n"
                            f"Sí = copiar a corregida (con nombre saneado y sufijo 'C')\n"
                            f"No = omitir este archivo"
                        )
                        if ans:
                            out_name_any = limit_filename(base_clean + SUFFIX_C + ext_lower, MAX_FILE_NAME_DEFAULT)
                            out_name_any = ensure_unique_preserving_C(dir_map[cur], out_name_any)
                            any_dir_final, out_name_any_final = fit_in_maxpath_bubbling(
                                target_dir=depth_ok_dir, floor_dir=base_dest, name=out_name_any, keep_C_suffix=True
                            )
                            any_target = any_dir_final / out_name_any_final
                            any_target.parent.mkdir(parents=True, exist_ok=True)
                            try:
                                estado = copy_with_prompt_on_timeout(src_file, any_target, relq, inner_cb=inner_cb)
                                if estado.startswith("COPIADO"):
                                    mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":str(any_target), "Estado":"COPIADO_STUCK_A_CORREGIDA"})
                                elif estado == "OMITIDO_STUCK":
                                    mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":"--OMITIDO (copia lenta)", "Estado":"OMITIDO_STUCK"})
                                else:
                                    mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":"--ERROR_STUCK_COPY", "Estado":"ERROR"})
                            except Exception as e:
                                mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":f"--ERROR_STUCK_COPY: {e}", "Estado":"ERROR"})
                        else:
                            mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":"--OMITIDO (stuck)", "Estado":"OMITIDO_STUCK"})
                    else:
                        rel_parts = list(src_file.relative_to(src_root).parts)
                        rel_sanit = [sanitize_component_letters_digits(p) for p in rel_parts[:-1]]
                        dump_subdir = dump_dir.joinpath(*rel_sanit) if rel_sanit else dump_dir
                        dump_subdir.mkdir(parents=True, exist_ok=True)

                        dump_subdir = bubble_file_for_maxdepth(dump_subdir, dump_dir)
                        dump_name = limit_filename(sanitize_component_letters_digits(src_file.stem) + ext_lower, MAX_FILE_NAME_DEFAULT)
                        dump_name = ensure_unique_generic(dump_subdir, dump_name)

                        dump_dir_final, dump_name_final = fit_in_maxpath_bubbling(
                            target_dir=dump_subdir, floor_dir=dump_dir, name=dump_name, keep_C_suffix=False
                        )
                        dump_target = dump_dir_final / dump_name_final
                        dump_target.parent.mkdir(parents=True, exist_ok=True)
                        try:
                            try:
                                relq = str(src_file.relative_to(src_root)).replace('\\', '/')
                            except Exception:
                                relq = src_file.name
                            estado = copy_with_prompt_on_timeout(src_file, dump_target, relq, inner_cb=inner_cb)
                            if estado.startswith("COPIADO"):
                                mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":str(dump_target), "Estado":"EXTRAIDO_NO_PDF"})
                            elif estado == "OMITIDO_STUCK":
                                mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":"--OMITIDO (dump lento)", "Estado":"OMITIDO_STUCK"})
                            else:
                                mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":"--ERROR_DUMP", "Estado":"ERROR"})
                        except Exception as e:
                            mapping.append({"Tipo":"FILE","Original":str(src_file), "Corregido":f"--ERROR_DUMP: {e}", "Estado":"ERROR"})

    if progress_cb:
        try: progress_cb(total, total, "Completado")
        except Exception: pass

    return mapping, base_dest, dump_dir

# ======= Reportes =======
def save_reports_with_label(results, counts, out_dir: Path, selected_root, label: str):
    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"{PROJECT_NAME} - {label}"
    html_path = out_dir / f"{base}.html"

    def filter_by(key): return [r for r in results if key in (r['Problemas'] or '')]
    forbidden_rows  = filter_by('Caracteres prohibidos') + filter_by('CaracteresProhibidos')
    diacritics_rows = filter_by('Diacríticos')
    path_rows       = filter_by('Ruta>MaxPath')
    name_rows       = filter_by('Nombre>MaxFileName')
    depth_rows      = filter_by('Profundidad>MaxDepth')
    hidden_rows     = filter_by('Oculto/Temporales')

    def render_table(title, rows, action):
        if not rows:
            return f"<h3>{title}</h3><p><em>Sin incidencias.</em></p>"
        body = "".join(
            f"<tr><td>{i}</td><td>{r['Tipo']}</td>"
            f"<td><code>{r['Ruta']}</code></td>"
            f"<td><code>{r['Nombre']}</code></td>"
            f"<td>{action}</td></tr>"
            for i, r in enumerate(rows, 1)
        )
        return (f"<h3>{title}</h3>"
                f"<table><thead><tr><th>#</th><th>Tipo</th><th>Ruta</th><th>Nombre</th><th>Corrección</th>"
                f"</tr></thead><tbody>{body}</tbody></table>")

    total = len(results)
    problematic = sum(1 for r in results if r['Problemas'])
    ok = total - problematic

    html = f"""<!doctype html>
<html lang="es"><head>
<meta charset="utf-8">
<title>{base} - Resumen</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 24px; }}
table {{ border-collapse: collapse; width: 100%; margin-top: 12px; }}
th, td {{ border: 1px solid #ddd; padding: 6px; font-size: 14px; }}
th {{ background: #f5f5f5; }}
.badge {{ display:inline-block; padding:4px 10px; border-radius:10px; font-weight:600; }}
.badge-ok {{ background:#c8f7c5; }}
.badge-bad {{ background:#ffd6d6; }}
code {{ background:#f1f1f1; padding:2px 4px; }}
</style></head><body>

<h1>{base} - Resumen de validación</h1>
<p><strong>Carpeta:</strong> <code>{selected_root}</code></p>

<div class="stats">
  <span class="total">Total elementos: <strong class="total">{total}</strong></span>
  &nbsp;|&nbsp; Sin problemas: <span class="badge badge-ok">{ok}</span>
  &nbsp;|&nbsp; Con problemas: <span class="badge badge-bad">{problematic}</span>
</div>

<h2>Parámetros</h2>
<ul>
  <li>Caracteres prohibidos (regex): <code>{FORBIDDEN_CHARS_PATTERN}</code></li>
  <li>Diacríticos (regex): <code>{DIACRITIC_CHARS_PATTERN}</code></li>
  <li>MaxPath: <code>{MAX_PATH_DEFAULT}</code></li>
  <li>MaxFileName: <code>{MAX_FILE_NAME_DEFAULT}</code></li>
  <li>MaxDepth: <code>{MAX_DEPTH_DEFAULT}</code></li>
  <li>Prefijo fusión carpetas: <code>{KEEP_SHORT_DIR}</code></li>
  <li>Acortado archivos: <code>{KEEP_SHORT_FILE}</code></li>
  <li>Sufijo aplicado: <code>{SUFFIX_C}</code></li>
  <li>NO-PDF extraídos a: <code>{NON_PDF_DUMP_NAME}</code></li>
</ul>

<h2>Detalles y acción propuesta</h2>
<div>{render_table("Caracteres prohibidos", forbidden_rows, "<strong>ELIMINAR</strong>")}</div>
<div>{render_table("Diacríticos", diacritics_rows, "<strong>ELIMINAR</strong>")}</div>
<div>{render_table("Ruta &gt; MaxPath", path_rows, "<strong>ACORTAR o SUBIR</strong>")}</div>
<div>{render_table("Nombre &gt; MaxFileName", name_rows, "<strong>RECORTAR</strong>")}</div>
<div>{render_table("Profundidad &gt; MaxDepth", depth_rows, "<strong>SUBIR AL PADRE</strong> (burbujeo)")}</div>
<div>{render_table("Ocultos / Temporales", hidden_rows, "<strong>OMITIR</strong> en copia")}</div>

</body></html>"""
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return None, html_path

def save_final_report(initial_counts: dict, corrected_counts: dict, mapping_rows: list,
                      out_dir: Path, selected_root: str, corrected_root: str,
                      initial_files_count: int, corrected_files_in_corr: int, extracted_nonpdf_count: int,
                      dump_path: str, mapping_csv_enabled: bool):
    base = f"{PROJECT_NAME} - FINAL"
    html_path = out_dir / f"{base}.html"

    total_files = sum(1 for m in mapping_rows if m.get("Tipo") == "FILE")
    conv = sum(1 for m in mapping_rows if m.get("Tipo") == "FILE" and m.get("Estado") == "CONVERTIDO")
    copi = sum(1 for m in mapping_rows if m.get("Tipo") == "FILE" and (m.get("Estado") in {"COPIADO","COPIADO_LENTO","COPIADO_STUCK_A_CORREGIDA"}))
    extracted = sum(1 for m in mapping_rows if m.get("Tipo") == "FILE" and m.get("Estado") == "EXTRAIDO_NO_PDF")
    err  = sum(1 for m in mapping_rows if m.get("Tipo") == "FILE" and ("ERROR" in m.get("Estado","")))
    omitted_stuck  = sum(1 for m in mapping_rows if m.get("Tipo") == "FILE" and m.get("Estado") == "OMITIDO_STUCK")
    omitted_hidden = sum(1 for m in mapping_rows if m.get("Tipo") == "FILE" and m.get("Estado") == "OMITIDO")
    omitted_total  = omitted_stuck + omitted_hidden

    same_count = (initial_files_count == (corrected_files_in_corr + extracted_nonpdf_count))
    status_badge = '<span class="badge badge-ok">OK</span>' if same_count else '<span class="badge badge-bad">MISMATCH</span>'

    size_original = dir_size_bytes(Path(selected_root))
    size_corr     = dir_size_bytes(Path(corrected_root))
    size_dump     = dir_size_bytes(Path(dump_path))
    size_final    = size_corr + size_dump
    delta_bytes   = size_final - size_original
    delta_pct     = (delta_bytes / size_original * 100.0) if size_original > 0 else 0.0

    def counts_table(title, c):
        return f"""
        <h3>{title}</h3>
        <table>
          <tr><th>Problema</th><th>Cantidad</th></tr>
          <tr><td>Caracteres prohibidos</td><td>{c['forbidden_chars']}</td></tr>
          <tr><td>Diacríticos</td><td>{c['diacritics']}</td></tr>
          <tr><td>Ruta &gt; MaxPath</td><td>{c['too_long_path']}</td></tr>
          <tr><td>Nombre &gt; MaxFileName</td><td>{c['too_long_name']}</td></tr>
          <tr><td>Profundidad &gt; MaxDepth</td><td>{c['too_deep']}</td></tr>
          <tr><td>Ocultos/Temporales</td><td>{c['hidden_temp']}</td></tr>
        </table>"""

    html = f"""<!doctype html>
<html lang="es"><head>
<meta charset="utf-8">
<title>{base}</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 24px; }}
table {{ border-collapse: collapse; width: 100%; margin-top: 12px; }}
th, td {{ border: 1px solid #ddd; padding: 6px; font-size: 14px; }}
th {{ background: #f5f5f5; }}
.badge {{ display:inline-block; padding:4px 10px; border-radius:10px; font-weight:600; }}
.badge-ok {{ background:#c8f7c5; }}
.badge-bad {{ background:#ffd6d6; }}
code {{ background:#f1f1f1; padding:2px 4px; }}
</style></head><body>

<h1>{PROJECT_NAME} - Informe FINAL</h1>
<p><strong>Original:</strong> <code>{selected_root}</code><br>
<strong>Corregido:</strong> <code>{corrected_root}</code><br>
<strong>Extraídos NO-PDF:</strong> <code>{dump_path}</code></p>

<h2>Integridad de cantidad de archivos {status_badge}</h2>
<table>
  <tr><th>Inicial</th><th>Corregida + Extraídos</th><th>Omitidos por 'stuck'</th></tr>
  <tr><td>{initial_files_count}</td><td>{corrected_files_in_corr + extracted_nonpdf_count}</td><td>{omitted_stuck}</td></tr>
</table>

<h2>Tamaño de carpetas</h2>
<table>
  <tr><th>Carpeta</th><th>Tamaño</th></tr>
  <tr><td>Original</td><td>{human_size(size_original)}</td></tr>
  <tr><td>Corregida</td><td>{human_size(size_corr)}</td></tr>
  <tr><td>Extraídos NO-PDF</td><td>{human_size(size_dump)}</td></tr>
  <tr><td><strong>Final (corregida + extraídos)</strong></td><td><strong>{human_size(size_final)}</strong></td></tr>
</table>
<p>Variación vs original: <strong>{'+' if delta_bytes>=0 else ''}{human_size(abs(delta_bytes))}</strong> ({delta_pct:+.1f}%).</p>

<h2>Resumen de procesamiento</h2>
<table>
  <tr><th>Total</th><th>Convertidos a PDF</th><th>Copiados</th><th>Extraídos NO-PDF</th><th>Omitidos por 'stuck'</th><th>Errores</th></tr>
  <tr><td>{total_files}</td><td>{conv}</td><td>{copi}</td><td>{extracted}</td><td>{omitted_stuck}</td><td>{err}</td></tr>
</table>

<h2>Archivos omitidos</h2>
<table>
  <tr><th>Tipo de omisión</th><th>Cantidad</th></tr>
  <tr><td>Omitidos (ocultos/temporales)</td><td>{omitted_hidden}</td></tr>
  <tr><td>Omitidos por 'stuck' (decisión del usuario)</td><td>{omitted_stuck}</td></tr>
  <tr><td><strong>Total omitidos</strong></td><td><strong>{omitted_total}</strong></td></tr>
</table>

<h2>Comparativo de incidencias</h2>
<div style="display:flex; gap:16px;">
  <div style="flex:1;">{counts_table("Antes (INICIAL)", initial_counts)}</div>
  <div style="flex:1;">{counts_table("Después (CORREGIDO)", corrected_counts)}</div>
</div>

</body></html>"""
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return html_path

# ======= GUI / Flujo =======
def run_gui():
    root = tk.Tk(); root.withdraw()
    messagebox.showinfo(PROJECT_NAME, "Selecciona la carpeta que deseas validar/corregir.")
    selected = filedialog.askdirectory(title="Selecciona la carpeta a validar")
    if not selected: return
    try:
        # INICIAL
        results_initial, counts_initial = validate_folder(selected)
        parent = Path(selected).resolve().parent
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = parent / REPORT_FOLDER_NAME / stamp
        out_dir.mkdir(parents=True, exist_ok=True)
        _, html_i = save_reports_with_label(results_initial, counts_initial, out_dir, selected, "INICIAL")

        initial_files_count = count_files(Path(selected))

        # Abrir INICIAL
        try:
            for _ in range(10):
                if Path(html_i).exists(): break
                time.sleep(0.2)
            try: webbrowser.open(Path(html_i).resolve().as_uri(), new=2)
            except Exception:
                try: os.startfile(html_i)  # type: ignore[attr-defined]
                except Exception:
                    subprocess.run(["explorer", "/select,", str(Path(html_i))], check=False)
        except Exception:
            pass

        if messagebox.askyesno(PROJECT_NAME, "¿Crear COPIA corregida (profundidad y path controlados) y convertir a PDF?"):
            # Ventana de progreso
            total_files_for_progress = max(initial_files_count, 1)
            prog = tk.Toplevel()
            prog.title("Progreso de corrección")
            prog.geometry("740x170")
            lbl = tk.Label(prog, text="Preparando...", anchor="w")
            lbl.pack(padx=12, pady=(12, 6), fill="x")
            bar = ttk.Progressbar(prog, orient="horizontal", mode="determinate", maximum=total_files_for_progress)
            bar.pack(padx=12, pady=(0, 8), fill="x")
            pct_lbl = tk.Label(prog, text="0%", anchor="e")
            pct_lbl.pack(padx=12, pady=(0, 8), fill="x")

            # Barra por archivo
            file_lbl = tk.Label(prog, text="Archivo: —", anchor="w")
            file_lbl.pack(padx=12, pady=(0, 2), fill="x")
            file_bar = ttk.Progressbar(prog, orient="horizontal", mode="determinate", maximum=100)
            file_bar.pack(padx=12, pady=(0, 4), fill="x")
            file_pct = tk.Label(prog, text="0% — 0 / 0 MB", anchor="e")
            file_pct.pack(padx=12, pady=(0, 12), fill="x")

            def progress_cb(current, total, rel_path_text):
                try:
                    bar.configure(maximum=max(total, 1))
                    bar['value'] = current
                    pct = int(current * 100 / total) if total else 100
                    lbl.configure(text=f"Procesando {current}/{total}: {rel_path_text}")
                    pct_lbl.configure(text=f"{pct}%")
                    _ui_pump()
                except Exception:
                    pass

            def file_progress_cb(done_bytes: int, total_bytes: int, name: str):
                try:
                    file_lbl.configure(text=f"Archivo: {name}")
                    if total_bytes > 0:
                        pct = int(done_bytes * 100 / total_bytes)
                    else:
                        pct = 0
                    file_bar['value'] = pct
                    file_pct.configure(
                        text=f"{pct}% — {human_size(done_bytes)} / {human_size(total_bytes)}"
                    )
                    _ui_pump()
                except Exception:
                    pass

            # Proceso con progreso
            mapping, corrected_root, dump_dir = copy_with_rules_and_convert(
                Path(selected), parent,
                progress_cb=progress_cb,
                total_files=initial_files_count,
                file_progress_cb=file_progress_cb
            )

            try: prog.destroy()
            except Exception: pass

            # Limpieza
            removed_corr = prune_empty_dirs(corrected_root, keep_root=True)
            removed_dump = prune_empty_dirs(dump_dir, keep_root=True)

            # CORREGIDO
            results_corr, counts_corr = validate_folder(str(corrected_root))
            _, html_c = save_reports_with_label(results_corr, counts_corr, out_dir, str(corrected_root), "CORREGIDO")

            corrected_files_in_corr = count_files(Path(corrected_root))
            extracted_nonpdf_count = count_files(Path(dump_dir))
            final_html = save_final_report(
                counts_initial, counts_corr, mapping, out_dir,
                selected, str(corrected_root),
                initial_files_count, corrected_files_in_corr, extracted_nonpdf_count,
                str(dump_dir), mapping_csv_enabled=GENERATE_MAPPING_CSV
            )

            final_total = corrected_files_in_corr + extracted_nonpdf_count
            if initial_files_count != final_total:
                messagebox.showwarning(PROJECT_NAME,
                    f"Atención: Cantidad de archivos distinta (posibles omitidos por 'stuck').\n"
                    f"Inicial: {initial_files_count}\nFinal (corregida + extraídos): {final_total}")

            messagebox.showinfo(PROJECT_NAME, f"✅ Proceso completado.\n\n"
                                              f"Carpetas vacías eliminadas -> Corregida: {removed_corr}, Dump: {removed_dump}\n\n"
                                              f"Reportes en:\n{out_dir}\n\n"
                                              f"- INICIAL: {html_i}\n- CORREGIDO: {html_c}\n- FINAL: {final_html}\n\n"
                                              f"Copia corregida:\n{corrected_root}\n\n"
                                              f"NO-PDF extraídos en:\n{dump_dir}")
            try: webbrowser.open(Path(final_html).resolve().as_uri(), new=2)
            except Exception:
                try: os.startfile(final_html)  # type: ignore[attr-defined]
                except Exception:
                    subprocess.run(["explorer", "/select,", str(Path(final_html))], check=False)
        else:
            messagebox.showinfo(PROJECT_NAME, f"Validación inicial generada.\n\nReportes en:\n{out_dir}\n\n{html_i}")
    except Exception as e:
        try:
            if tk._default_root:
                for w in tk._default_root.winfo_children():
                    if isinstance(w, tk.Toplevel):
                        w.destroy()
        except Exception:
            pass
        messagebox.showerror(PROJECT_NAME, f"Ocurrió un error:\n{e}")

# ======= Builder integrado (PyInstaller onefile; coloca tcl/tk en _tk_data) =======
def _build_onefile():
    """
    Compila este mismo script a UN solo .exe y coloca Tcl/Tk donde lo espera
    el hook oficial de PyInstaller: _tk_data/tcl8.x y _tk_data/tk8.x
    Uso:
      python este_script.py --build --name "SECOPII-DOC-CHECK" --onefile --windowed [--icon icon.ico]
    """
    import argparse
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("--build", action="store_true", help="Generar .exe y salir")
    ap.add_argument("--name", default="SECOPII-DOC-CHECK", help="Nombre del .exe")
    ap.add_argument("--onefile", action="store_true", help="Empaquetar en un solo archivo")
    ap.add_argument("--windowed", action="store_true", help="Sin consola (GUI)")
    ap.add_argument("--console", action="store_true", help="Consola visible (anula --windowed)")
    ap.add_argument("--icon", default=None, help="Ruta a .ico opcional")
    args, _unknown = ap.parse_known_args()

    if not args.build:
        return False

    try:
        from PyInstaller.__main__ import run as pyirun
    except Exception:
        print("PyInstaller no está instalado. Ejecuta:\n  python -m pip install --upgrade pyinstaller")
        sys.exit(1)

    # Detectar Tcl/Tk de esta instalación de Python
    tcl_root = os.path.join(sys.exec_prefix, "tcl")
    cand_tcl = sorted(glob.glob(os.path.join(tcl_root, "tcl8.*")), reverse=True)
    cand_tk  = sorted(glob.glob(os.path.join(tcl_root, "tk8.*")),  reverse=True)
    tcl_dir = cand_tcl[0] if cand_tcl else None
    tk_dir  = cand_tk[0]  if cand_tk  else None

    # Colocar datos donde los espera pyi_rth_tkinter: _tk_data/<tcl|tk>8.x
    sep = ";" if os.name == "nt" else ":"
    add_data = []
    if tcl_dir:
        add_data += ["--add-data", f"{tcl_dir}{sep}_tk_data/{os.path.basename(tcl_dir)}"]
    if tk_dir:
        add_data += ["--add-data", f"{tk_dir}{sep}_tk_data/{os.path.basename(tk_dir)}"]

    # Hidden imports necesarios
    hidden = ["tkinter", "_tkinter"]
    if importlib.util.find_spec("docx2pdf"):
        hidden += ["docx2pdf", "comtypes", "comtypes.client", "win32com"]
    if importlib.util.find_spec("PIL"):
        hidden += ["PIL", "PIL.Image", "PIL._imaging"]
    hidden_args = []
    for m in hidden:
        hidden_args += ["--hidden-import", m]

    # Argumentos PyInstaller
    argv = [
        "--noconfirm", "--clean",
        "--name", args.name,
        __file__,
        *add_data,
        *hidden_args,
    ]
    if args.onefile or True:    # forzamos onefile por tu requerimiento
        argv.insert(0, "--onefile")
    if args.windowed and not args.console:
        argv.insert(0, "--windowed")
    if args.icon:
        argv += ["--icon", args.icon]

    print("\n[BUILD] PyInstaller:", " ".join(argv), "\n")
    pyirun(argv)
    print(f"\n[BUILD] Listo: .\\dist\\{args.name}.exe")
    return True

# ======= Main =======
if __name__ == "__main__":
    # Si se llamó con --build, compila y sale; si no, corre la app normal.
    if _build_onefile():
        sys.exit(0)
    run_gui()

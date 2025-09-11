# -*- coding: utf-8 -*-
"""
SECOP II DOC CHECK
- Selecciona carpeta (GUI)
- Valida: caracteres prohibidos, diacríticos, longitud de ruta/nombre y profundidad
- Reportes: parent(selected)/SECOP_DOC_CHECK/AAAAMMDD_HHMMSS/
- Archivos: "SECOP II DOC CHECK.html" y ".csv"
"""

import os
import re
import csv
import platform
import webbrowser
import subprocess
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

# ======= Configuración =======
PROJECT_NAME = "SECOP II DOC CHECK"
REPORT_FOLDER_NAME = "SECOP_DOC_CHECK"

# Límites
MAX_PATH_DEFAULT = 240
MAX_FILE_NAME_DEFAULT = 100
MAX_DEPTH_DEFAULT = 5

# Reglas
FORBIDDEN_CHARS_PATTERN = r'[\\/:*?"<>|%&#+\{\}\[\];,=]'
DIACRITIC_CHARS_PATTERN  = r'[áéíóúüÁÉÍÓÚÜñÑçÇãÃõÕ]'
HIDDEN_BASENAMES = {'Thumbs.db', '.DS_Store', '.ds_store', 'desktop.ini'}

# ======= Utilidades =======
def rel_depth(full_path: str, root: str) -> int:
    rel = os.path.relpath(full_path, root)
    if rel in ('.', ''):
        return 0
    return len([p for p in re.split(r'[\\/]+', rel) if p])

def safe_walk(root: str):
    """os.walk tolerante a rutas largas en Windows."""
    if platform.system() == 'Windows':
        def add_prefix(p):
            p = os.path.abspath(p)
            if p.startswith('\\\\?\\') or p.startswith('\\\\.\\'):
                return p
            if p.startswith('\\\\'):
                return '\\\\?\\UNC\\' + p[2:]
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

# ======= Validación =======
def _validate_item(name, full, tipo, root, max_path, max_file_name, max_depth,
                   forbidden_re, diacritics_re, counts):
    plen, nlen, depth = len(full), len(name), rel_depth(full, root)
    issues = []

    if forbidden_re.search(name):
        issues.append('CaracteresProhibidos'); counts['forbidden_chars'] += 1
    if diacritics_re.search(name):
        issues.append('Diacriticos'); counts['diacritics'] += 1
    if plen > max_path:
        issues.append('Ruta>MaxPath'); counts['too_long_path'] += 1
    if nlen > max_file_name:
        issues.append('Nombre>MaxFileName'); counts['too_long_name'] += 1
    if depth > max_depth:
        issues.append('Profundidad>MaxDepth'); counts['too_deep'] += 1
    if name in HIDDEN_BASENAMES or name.startswith('~$'):
        issues.append('Oculto/Temporal'); counts['hidden_temp'] += 1

    order = {'CaracteresProhibidos': 0, 'Diacriticos': 1,
             'Ruta>MaxPath': 2, 'Nombre>MaxFileName': 3,
             'Profundidad>MaxDepth': 4, 'Oculto/Temporal': 5}
    issues.sort(key=lambda x: order.get(x, 99))

    return {
        'Tipo': tipo,
        'Ruta': full,
        'Nombre': name,
        'LongRuta': plen,
        'LongNombre': nlen,
        'Profundidad': depth,
        'Problemas': ','.join(issues)
    }

def validate_folder(source_folder,
                    max_path=MAX_PATH_DEFAULT,
                    max_file_name=MAX_FILE_NAME_DEFAULT,
                    max_depth=MAX_DEPTH_DEFAULT):
    results = []
    counts = {'too_long_path': 0, 'too_long_name': 0, 'too_deep': 0,
              'forbidden_chars': 0, 'diacritics': 0, 'hidden_temp': 0}

    forbidden_re = re.compile(FORBIDDEN_CHARS_PATTERN)
    diacritics_re = re.compile(DIACRITIC_CHARS_PATTERN, re.UNICODE)
    source_folder = os.path.abspath(source_folder)

    for dirpath, dirnames, filenames in safe_walk(source_folder):
        current_dir_name = os.path.basename(dirpath)
        if current_dir_name:
            results.append(_validate_item(current_dir_name, dirpath, 'DIR',
                                          source_folder, max_path, max_file_name, max_depth,
                                          forbidden_re, diacritics_re, counts))
        for d in dirnames:
            results.append(_validate_item(d, os.path.join(dirpath, d), 'DIR',
                                          source_folder, max_path, max_file_name, max_depth,
                                          forbidden_re, diacritics_re, counts))
        for fn in filenames:
            results.append(_validate_item(fn, os.path.join(dirpath, fn), 'FILE',
                                          source_folder, max_path, max_file_name, max_depth,
                                          forbidden_re, diacritics_re, counts))

    def prio(row):
        probs = row.get('Problemas') or ''
        keys = ['CaracteresProhibidos','Diacriticos','Ruta>MaxPath',
                'Nombre>MaxFileName','Profundidad>MaxDepth','Oculto/Temporal']
        idx = min([keys.index(k) for k in keys if k in probs] or [99])
        return (idx, -row.get('LongRuta', 0))
    results.sort(key=prio)
    return results, counts

# ======= Reportes =======
def save_reports(results, counts, out_dir: Path,
                 max_path, max_file_name, max_depth, selected_root):
    out_dir.mkdir(parents=True, exist_ok=True)

    csv_path = out_dir / f"{PROJECT_NAME}.csv"
    html_path = out_dir / f"{PROJECT_NAME}.html"

    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'Tipo','Ruta','Nombre','LongRuta','LongNombre','Profundidad','Problemas'
        ])
        writer.writeheader()
        writer.writerows(results)

    def filter_by(key): return [r for r in results if key in (r['Problemas'] or '')]
    forbidden_rows  = filter_by('CaracteresProhibidos')
    diacritics_rows = filter_by('Diacriticos')
    path_rows       = filter_by('Ruta>MaxPath')
    name_rows       = filter_by('Nombre>MaxFileName')
    depth_rows      = filter_by('Profundidad>MaxDepth')
    hidden_rows     = filter_by('Oculto/Temporal')

    def render_table(title, rows):
        if not rows:
            return f"<h3>{title}</h3><p><em>Sin incidencias.</em></p>"
        body = "".join(
            f"<tr><td>{i}</td><td>{r['Tipo']}</td>"
            f"<td><code>{r['Ruta']}</code></td>"
            f"<td><code>{r['Nombre']}</code></td>"
            f"<td>{r['Problemas']}</td></tr>"
            for i, r in enumerate(rows, 1)
        )
        return (
            f"<h3>{title}</h3>"
            f"<table><thead><tr>"
            f"<th>#</th><th>Tipo</th><th>Ruta</th><th>Nombre</th><th>Problemas</th>"
            f"</tr></thead><tbody>{body}</tbody></table>"
        )

    total = len(results)
    problematic = sum(1 for r in results if r['Problemas'])
    ok = total - problematic

    html = f"""<!doctype html>
<html lang="es"><head>
<meta charset="utf-8">
<title>{PROJECT_NAME} - Resumen</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 24px; }}
table {{ border-collapse: collapse; width: 100%; margin-top: 12px; }}
th, td {{ border: 1px solid #ddd; padding: 6px; font-size: 14px; }}
th {{ background: #f5f5f5; }}
.badge {{ display:inline-block; padding:4px 10px; border-radius:10px; font-weight:600; }}
.badge-ok {{ background:#c8f7c5; }}
.badge-bad {{ background:#ffd6d6; }}
.stats {{ margin: 10px 0 18px 0; font-size: 18px; }}
.stats strong.total {{ font-size: 22px; }}
code {{ background:#f1f1f1; padding:2px 4px; border-radius:4px; }}
</style></head><body>

<h1>{PROJECT_NAME} - Resumen de validación</h1>
<p><strong>Carpeta seleccionada:</strong> <code>{selected_root}</code></p>

<div class="stats">
  <span class="total">Total elementos: <strong class="total">{total}</strong></span>
  &nbsp;|&nbsp; Sin problemas: <span class="badge badge-ok">{ok}</span>
  &nbsp;|&nbsp; Con problemas: <span class="badge badge-bad">{problematic}</span>
</div>

<h2>Parámetros</h2>
<ul>
  <li>Caracteres prohibidos (regex): <code>{FORBIDDEN_CHARS_PATTERN}</code></li>
  <li>Diacríticos (regex): <code>{DIACRITIC_CHARS_PATTERN}</code></li>
  <li>MaxPath: <code>{max_path}</code></li>
  <li>MaxFileName: <code>{max_file_name}</code></li>
  <li>MaxDepth: <code>{max_depth}</code></li>
</ul>

<h2>Conteo general</h2>
<table>
  <tr><th>Problema</th><th>Cantidad</th></tr>
  <tr><td>Caracteres prohibidos</td><td>{counts['forbidden_chars']}</td></tr>
  <tr><td>Diacríticos</td><td>{counts['diacritics']}</td></tr>
  <tr><td>Ruta &gt; MaxPath</td><td>{counts['too_long_path']}</td></tr>
  <tr><td>Nombre &gt; MaxFileName</td><td>{counts['too_long_name']}</td></tr>
  <tr><td>Profundidad &gt; MaxDepth</td><td>{counts['too_deep']}</td></tr>
  <tr><td>Ocultos/Temporales</td><td>{counts['hidden_temp']}</td></tr>
</table>

<h2>Detalles</h2>
<div>{render_table("Caracteres prohibidos", forbidden_rows)}</div>
<div>{render_table("Diacríticos", diacritics_rows)}</div>
<div>{render_table("Ruta > MaxPath", path_rows)}</div>
<div>{render_table("Nombre > MaxFileName", name_rows)}</div>
<div>{render_table("Profundidad > MaxDepth", depth_rows)}</div>
<div>{render_table("Ocultos / Temporales", hidden_rows)}</div>

</body></html>"""
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)

    return csv_path, html_path

# ======= App (GUI) =======
def run_gui():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(PROJECT_NAME, "Selecciona la carpeta que deseas validar.")
    selected = filedialog.askdirectory(title="Selecciona la carpeta a validar")
    if not selected:
        return
    try:
        results, counts = validate_folder(selected)

        parent = Path(selected).resolve().parent
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = parent / REPORT_FOLDER_NAME / stamp  # ruta más corta

        csv_path, html_path = save_reports(
            results, counts, out_dir,
            MAX_PATH_DEFAULT, MAX_FILE_NAME_DEFAULT, MAX_DEPTH_DEFAULT, selected
        )

        # --- Apertura robusta del HTML ---
        try:
            for _ in range(10):
                if Path(html_path).exists():
                    break
                time.sleep(0.2)

            opened = False
            try:
                opened = webbrowser.open(Path(html_path).resolve().as_uri(), new=2)
            except Exception:
                opened = False

            if not opened:
                try:
                    os.startfile(html_path)  # type: ignore[attr-defined]
                    opened = True
                except Exception:
                    opened = False

            if not opened:
                try:
                    subprocess.run(["explorer", "/select,", str(Path(html_path))], check=False)
                except Exception:
                    pass

        except Exception as e:
            messagebox.showwarning(PROJECT_NAME, f"No se pudo abrir el HTML automáticamente.\nRuta: {html_path}\n\n{e}")

        messagebox.showinfo(PROJECT_NAME, f"Validación completada.\n\nCSV:\n{csv_path}\nHTML:\n{html_path}")

    except Exception as e:
        messagebox.showerror(PROJECT_NAME, f"Ocurrió un error:\n{e}")

if __name__ == "__main__":
    run_gui()
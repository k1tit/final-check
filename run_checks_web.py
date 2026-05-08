# -*- coding: utf-8 -*-
from __future__ import annotations
import argparse
import json
import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BUILD_SCRIPT = ROOT / "build_checks.py"
STATUS_FILE = ROOT / "status.json"
BUILD_LOG = ROOT / "last_build_checks.log"
BUILD_PROGRESS_FILE = ROOT / "build_progress.json"
RUNTIME_PATHS_JSON = ROOT / "runtime_paths.json"
DATA_DIR = ROOT / "data"
DEFAULT_WEB_EXCEPTION_XLSX = DATA_DIR / "Exception.xlsx"
STARTUP_LOG = ROOT / "run_checks_web_startup.log"
DEFAULT_BASE_DIR = Path(r"C:\Users\kitit\check\check_PF_BP-PY-ZY\data\1 Нулевые файлы выгрузки макроса + файл исключений")
DEFAULT_OUTPUT_DIR = Path(r"C:\Users\kitit\check\check_PF_BP-PY-ZY\data\result")

_run_lock = threading.Lock()


def _append_startup_log(message: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(STARTUP_LOG, "a", encoding="utf-8", errors="replace") as f:
        f.write(f"[{ts}] {message}\n")


def _normalize_path_text(raw: str) -> str:
    txt = str(raw or "").strip().strip('"').strip("'")
    if not txt:
        return ""
    txt = os.path.expandvars(os.path.expanduser(txt))
    return str(Path(txt))


def _load_runtime_paths() -> dict[str, Path]:
    base_dir = DEFAULT_BASE_DIR
    output_dir = DEFAULT_OUTPUT_DIR
    exception_file = DEFAULT_WEB_EXCEPTION_XLSX
    try:
        if RUNTIME_PATHS_JSON.is_file():
            cfg = json.loads(RUNTIME_PATHS_JSON.read_text(encoding="utf-8"))
            base_raw = str(cfg.get("base_dir", "")).strip()
            out_raw = str(cfg.get("output_dir", "")).strip()
            exc_raw = str(cfg.get("exception_file", "")).strip()
            if base_raw:
                base_dir = Path(base_raw)
            if out_raw:
                output_dir = Path(out_raw)
            if exc_raw:
                exception_file = Path(exc_raw)
    except (OSError, ValueError, TypeError):
        pass
    return {"base_dir": base_dir, "output_dir": output_dir, "exception_file": exception_file}


def _save_runtime_paths(base_dir: str, output_dir: str, exception_file: str) -> dict[str, str]:
    b = Path(_normalize_path_text(base_dir))
    o = Path(_normalize_path_text(output_dir))
    e = Path(_normalize_path_text(exception_file))
    payload = {"base_dir": str(b), "output_dir": str(o), "exception_file": str(e)}
    tmp = RUNTIME_PATHS_JSON.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(RUNTIME_PATHS_JSON)
    return payload

def _probe_port_available(host: str, port: int) -> None:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        s.bind((host, port))
    except OSError as exc:
        raise SystemExit(f"Порт {port} уже занят. {exc}") from exc
    finally:
        s.close()


def _resolve_python_executable() -> str:
    candidates: list[str] = []
    base_exe = getattr(sys, "_base_executable", "")
    if base_exe:
        candidates.append(base_exe)
    candidates.append(sys.executable)
    py_in_path = shutil.which("python")
    if py_in_path:
        candidates.append(py_in_path)
    for raw in candidates:
        if not raw:
            continue
        p = Path(raw)
        name = p.name.lower()
        if name == "pythonservice.exe":
            for cand in (p.with_name("python.exe"), p.with_name("pythonw.exe")):
                if cand.exists():
                    return str(cand)
            continue
        if p.exists():
            return str(p)
    return str(Path(sys.executable))


def create_app():
    try:
        from flask import Flask, request, Response, jsonify
    except ImportError as exc:
        raise SystemExit("Установите: pip install flask") from exc

    app = Flask(__name__)

    def exception_file_path() -> Path:
        return _load_runtime_paths()["exception_file"]

    def _norm(v: object) -> str:
        s = str(v if v is not None else "").strip()
        if s.endswith(".0"):
            s = s[:-2]
        return s

    def _load_exceptions_df():
        import pandas as pd
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        exc_file = exception_file_path()
        if exc_file.exists():
            df = pd.read_excel(exc_file)
        else:
            df = pd.DataFrame(columns=["SO", "Customer", "Comment OM"])
        for col in ("SO", "Customer", "Comment OM"):
            if col not in df.columns:
                df[col] = ""
        return df[["SO", "Customer", "Comment OM"]]

    @app.route("/")
    def index():
        def esc_path(p: Path) -> str:
            return str(p.resolve()).replace("&", "&amp;").replace("<", "&lt;")

        paths = _load_runtime_paths()
        exc_path_disp = esc_path(paths["exception_file"])
        base_dir_disp = esc_path(paths["base_dir"])
        output_dir_disp = esc_path(paths["output_dir"])
        html = """<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="utf-8">
    <title>Проверка PF BP-PY-ZY</title>
    <style>
        body { font-family: system-ui, sans-serif; max-width: 50rem; margin: 2rem auto; padding: 0 1rem; }
        h1 { font-size: 1.5rem; margin-bottom: 0.5rem; }
        .info { background: #e3f2fd; padding: 1rem; border-radius: 8px; margin: 1rem 0; }
        label { display: block; margin-top: 1rem; font-weight: 600; }
        select, input { width: 100%; max-width: 30rem; padding: 0.5rem; margin-top: 0.25rem; }
        .checkbox-label { display: flex; align-items: center; gap: 0.5rem; margin-top: 1rem; font-weight: normal; }
        .checkbox-label input { width: auto; margin-top: 0; }
        button { margin-top: 1.5rem; padding: 0.5rem 1.5rem; font-size: 1rem; cursor: pointer; background: #007bff; color: white; border: none; border-radius: 4px; }
        button:hover { background: #0056b3; }
        button:disabled { opacity: 0.5; cursor: not-allowed; }
        .progress { margin: 1rem 0; display: none; }
        .bar { width: 100%; background: #e0e0e0; border-radius: 4px; overflow: hidden; }
        .fill { width: 0%; height: 30px; background: #28a745; transition: width 0.35s ease-out; text-align: center; color: white; line-height: 30px; }
        .fill.fill-running { background: linear-gradient(90deg, #0d6efd, #198754); }
        .message { margin-top: 0.5rem; color: #555; white-space: pre-line; }
        .result { background: #d4edda; color: #155724; padding: 1rem; border-radius: 8px; margin: 1rem 0; display: none; }
        .result-error { background: #f8d7da; color: #721c24; padding: 1rem; border-radius: 8px; margin: 1rem 0; display: none; white-space: pre-wrap; font-family: ui-monospace, monospace; font-size: 0.85rem; max-height: 20rem; overflow: auto; }
        .exceptions-panel { background: #fff3cd; border: 1px solid #ffecb5; border-radius: 8px; padding: 1rem; margin: 1rem 0; }
        .exceptions-panel h3 { margin: 0 0 0.5rem 0; font-size: 1rem; }
        .exceptions-list { margin: 0.5rem 0; max-height: 360px; overflow-y: auto; font-size: 0.9rem; }
        .exception-item { background: white; padding: 0.5rem; margin: 0.25rem 0; border-radius: 4px; display: flex; justify-content: space-between; align-items: center; border: 1px solid #dee2e6; }
        .exception-item button { margin: 0; padding: 0.25rem 0.5rem; font-size: 0.8rem; background: #dc3545; }
        .add-exception { display: flex; gap: 0.5rem; margin-top: 0.5rem; flex-wrap: wrap; }
        .add-exception input { flex: 1; margin-top: 0; min-width: 120px; }
        .add-exception button { margin-top: 0; padding: 0.5rem 1rem; background: #28a745; }
        .save-btn { background: #17a2b8; margin-top: 0.5rem; width: auto; }
        .secondary-btn { background: #6c757d; margin-left: 0.5rem; }
        .path-row { margin-top: 0.5rem; }
        .path-row input { max-width: none; width: 100%; }
    </style>
</head>
<body>
    <h1>Проверка PF BP-PY-ZY</h1>

    <div class="exceptions-panel">
        <h3>Пути сервиса</h3>
        <label for="baseDirPath">Путь к нулевым выгрузкам</label>
        <div class="path-row"><input type="text" id="baseDirPath" value="__BASE_DIR__" placeholder="Путь к нулевым выгрузкам"></div>
        <label for="outputDirPath">Путь к результатам</label>
        <div class="path-row"><input type="text" id="outputDirPath" value="__OUTPUT_DIR__" placeholder="Путь к результатам"></div>
        <label for="exceptionFilePath">Файл исключений (Exception.xlsx)</label>
        <div class="path-row"><input type="text" id="exceptionFilePath" value="__EXC_PATH__" placeholder="Путь к Exception.xlsx"></div>
        <button type="button" id="savePathsBtn">Сохранить пути</button>
        <p style="margin:0.5rem 0 0 0; font-size:0.9rem;">После сохранения новые пути применяются к следующему запуску формирования.</p>
    </div>

    <form id="runForm">
        <label for="mode">Режим</label>
        <select name="mode" id="mode" required>
            <option value="pairs">Три пары (3801+3803, 3802+3804, 3805+3806)</option>
            <option value="single">Каждая организация отдельно</option>
            <option value="custom_single">Несколько выбранных SO — по одной задаче</option>
            <option value="custom_group">Несколько выбранных SO — одной группой</option>
        </select>

        <label for="folders">Номера папок SO через запятую</label>
        <input type="text" name="folders" id="folders" placeholder="3801, 3803">

        <label class="checkbox-label">
            <input type="checkbox" name="skip_manual" id="skipManual" value="1" checked>
            <span>Не спрашивать дополнительные Exception при запуске скрипта (веб всегда без запросов)</span>
        </label>

        <div id="exceptionsPanel" class="exceptions-panel">
            <h3>Дополнительные исключения</h3>
            <p style="margin:0 0 0.5rem 0; font-size:0.9rem;">Поля SO и Customer: можно добавить сколько угодно строк — «+ Добавить». Файл: data\\Exception.xlsx</p>
            <div style="font-size:0.9rem; margin-bottom: 0.35rem;"><strong>Новые к добавлению:</strong></div>
            <div id="exceptionsList" class="exceptions-list"></div>
            <div class="add-exception">
                <input type="text" id="newSO" placeholder="SO (например: 3801)" autocomplete="off">
                <input type="text" id="newCustomer" placeholder="Customer" autocomplete="off">
                <button type="button" id="addExceptionBtn">+ Добавить</button>
            </div>
            <button type="button" id="saveExceptionsBtn" class="save-btn">Сохранить в data\\Exception.xlsx</button>
            <div style="font-size:0.9rem; margin: 0.7rem 0 0.35rem 0;"><strong>К удалению:</strong></div>
            <div id="deleteExceptionsList" class="exceptions-list"></div>
            <div class="add-exception">
                <input type="text" id="delSO" placeholder="SO (например: 3801)" autocomplete="off">
                <input type="text" id="delCustomer" placeholder="Customer" autocomplete="off">
                <button type="button" id="addDeleteExceptionBtn">+ Добавить к удалению</button>
            </div>
            <button type="button" id="deleteExceptionsBtn" class="save-btn secondary-btn">Удалить из файла</button>
        </div>

        <button type="submit" id="submitBtn">Сформировать отчёт</button>
    </form>

    <div id="progress" class="progress">
        <div class="bar"><div id="fill" class="fill">0%</div></div>
        <div id="msg" class="message"></div>
    </div>

    <div id="result" class="result">✓ Отчёт сформирован!</div>
    <div id="resultError" class="result-error"></div>

    <script>
        let exceptions = [];
        let deleteExceptions = [];
        
        function renderExceptions() {
            const container = document.getElementById('exceptionsList');
            if (exceptions.length === 0) {
                container.innerHTML = '<em>Нет добавленных исключений</em>';
                return;
            }
            let html = '';
            exceptions.forEach((exc, idx) => {
                html += `<div class="exception-item">
                    <span>SO: ${exc.so} | Customer: ${exc.customer}</span>
                    <button onclick="exceptions.splice(${idx},1); renderExceptions();">Удалить</button>
                </div>`;
            });
            container.innerHTML = html;
        }

        function renderDeleteExceptions() {
            const container = document.getElementById('deleteExceptionsList');
            if (deleteExceptions.length === 0) {
                container.innerHTML = '<em>Нет добавленных к удалению</em>';
                return;
            }
            let html = '';
            deleteExceptions.forEach((exc, idx) => {
                html += `<div class="exception-item">
                    <span>SO: ${exc.so} | Customer: ${exc.customer}</span>
                    <button onclick="deleteExceptions.splice(${idx},1); renderDeleteExceptions();">Убрать</button>
                </div>`;
            });
            container.innerHTML = html;
        }

        document.getElementById('addDeleteExceptionBtn').onclick = function() {
            const so = document.getElementById('delSO').value.trim();
            const cust = document.getElementById('delCustomer').value.trim();
            if (so && cust) {
                deleteExceptions.push({so: so, customer: cust});
                renderDeleteExceptions();
                document.getElementById('delSO').value = '';
                document.getElementById('delCustomer').value = '';
            } else {
                alert('Заполните оба поля для удаления');
            }
        };
        document.getElementById('deleteExceptionsBtn').onclick = async function() {
            if (deleteExceptions.length === 0) {
                alert('Нет исключений для удаления');
                return;
            }
            const resp = await fetch('/exceptions/delete', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({exceptions: deleteExceptions})
            });
            const result = await resp.json();
            if (!result.success) {
                alert('Ошибка удаления: ' + (result.error || 'unknown'));
                return;
            }
            alert(`Удалено: ${result.deleted}`);
            deleteExceptions = [];
            renderDeleteExceptions();
        };
        document.getElementById('savePathsBtn').onclick = async function() {
            const base_dir = document.getElementById('baseDirPath').value.trim();
            const output_dir = document.getElementById('outputDirPath').value.trim();
            const exception_file = document.getElementById('exceptionFilePath').value.trim();
            if (!base_dir || !output_dir || !exception_file) {
                alert('Заполните все пути');
                return;
            }
            const resp = await fetch('/paths', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({base_dir, output_dir, exception_file})
            });
            const result = await resp.json();
            if (result.success) {
                alert('Пути сохранены');
            } else {
                alert('Ошибка сохранения путей: ' + (result.error || 'unknown'));
            }
        };
        
        document.getElementById('addExceptionBtn').onclick = function() {
            const so = document.getElementById('newSO').value.trim();
            const cust = document.getElementById('newCustomer').value.trim();
            if (so && cust) {
                exceptions.push({so: so, customer: cust});
                renderExceptions();
                document.getElementById('newSO').value = '';
                document.getElementById('newCustomer').value = '';
            } else {
                alert('Заполните оба поля');
            }
        };
        
        document.getElementById('saveExceptionsBtn').onclick = async function() {
            if (exceptions.length === 0) {
                alert('Нет исключений для сохранения');
                return;
            }
            const resp = await fetch('/save_exceptions', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({exceptions: exceptions})
            });
            const result = await resp.json();
            if (result.success) {
                alert(`Сохранено ${result.saved} исключений`);
                exceptions = [];
                renderExceptions();
            } else {
                alert('Ошибка: ' + result.error);
            }
        };
        renderDeleteExceptions();
        
        document.getElementById('runForm').onsubmit = async function(e) {
            e.preventDefault();
            const btn = document.getElementById('submitBtn');
            const progress = document.getElementById('progress');
            btn.disabled = true;
            progress.style.display = 'block';
            document.getElementById('result').style.display = 'none';
            document.getElementById('resultError').style.display = 'none';
            
            if (exceptions.length > 0) {
                await document.getElementById('saveExceptionsBtn').onclick();
            }
            
            const formData = new FormData(this);
            await fetch('/run', {method: 'POST', body: formData});
            
            const interval = setInterval(async () => {
                const resp = await fetch('/status');
                const data = await resp.json();
                const fill = document.getElementById('fill');
                fill.style.width = data.percent + '%';
                fill.textContent = data.percent + '%';
                if (data.done) { fill.classList.remove('fill-running'); }
                else { fill.classList.add('fill-running'); }
                const msgEl = document.getElementById('msg');
                if (data.done) {
                    msgEl.textContent = data.message || '';
                } else {
                    msgEl.textContent = (data.message || '') + (data.detail ? ('\\n' + data.detail) : '');
                }
                if (data.done) {
                    clearInterval(interval);
                    btn.disabled = false;
                    const ok = data.ok !== false;
                    if (ok) {
                        document.getElementById('result').style.display = 'block';
                        document.getElementById('resultError').style.display = 'none';
                    } else {
                        document.getElementById('result').style.display = 'none';
                        const err = document.getElementById('resultError');
                        err.style.display = 'block';
                        err.textContent = (data.message || 'Ошибка') + (data.detail ? ('\\n\\n' + data.detail) : '');
                    }
                    setTimeout(() => progress.style.display = 'none', ok ? 3000 : 8000);
                }
            }, 500);
        };
    </script>
</body>
</html>"""
        return (
            html.replace("__EXC_PATH__", exc_path_disp)
            .replace("__BASE_DIR__", base_dir_disp)
            .replace("__OUTPUT_DIR__", output_dir_disp)
        )

    @app.route("/run", methods=["POST"])
    def run_checks():
        mode = request.form.get("mode", "pairs")
        folders = request.form.get("folders", "")
        skip_manual = request.form.get("skip_manual") == "1"

        def run():
            with _run_lock:
                cmd = [_resolve_python_executable(), str(BUILD_SCRIPT), "--mode", mode, "--skip-manual-exceptions"]
                if folders:
                    cmd.extend(["--folders", folders])
                paths_cfg = _load_runtime_paths()

                ok = False
                detail = ""
                proc: subprocess.CompletedProcess[str] | None = None
                child_env = {
                    **os.environ,
                    "PYTHONUNBUFFERED": "1",
                    "REPORTS_BASE_DIR": str(paths_cfg["base_dir"]),
                    "REPORTS_OUTPUT_DIR": str(paths_cfg["output_dir"]),
                    "REPORTS_EXCEPTION_FILE": str(paths_cfg["exception_file"]),
                }
                try:
                    if not paths_cfg["base_dir"].exists():
                        detail = f"Каталог нулевых выгрузок не найден: {paths_cfg['base_dir']}"
                        with open(BUILD_LOG, "w", encoding="utf-8", errors="replace") as logf:
                            logf.write(f"[runner] {detail}\n")
                        with open(STATUS_FILE, "w", encoding="utf-8") as f:
                            json.dump(
                                {
                                    "done": True,
                                    "ok": False,
                                    "percent": 100,
                                    "message": "Формирование отчёта завершилось с ошибкой",
                                    "detail": detail,
                                },
                                f,
                                ensure_ascii=False,
                            )
                        return
                    paths_cfg["output_dir"].mkdir(parents=True, exist_ok=True)
                    paths_cfg["exception_file"].parent.mkdir(parents=True, exist_ok=True)
                    with open(BUILD_LOG, "w", encoding="utf-8", errors="replace") as logf:
                        logf.write(f"[runner] cmd: {' '.join(cmd)}\n")
                        logf.write(f"[runner] base_dir: {paths_cfg['base_dir']}\n")
                        logf.write(f"[runner] output_dir: {paths_cfg['output_dir']}\n")
                        logf.write(f"[runner] exception_file: {paths_cfg['exception_file']}\n")
                        logf.flush()
                        proc = subprocess.run(
                            cmd,
                            cwd=str(ROOT),
                            stdout=logf,
                            stderr=subprocess.STDOUT,
                            shell=False,
                            text=True,
                            env=child_env,
                        )
                    code = proc.returncode
                    ok = code == 0
                    if not ok:
                        try:
                            log_text = BUILD_LOG.read_text(encoding="utf-8", errors="replace").strip()
                            lines = log_text.splitlines()
                            tail = "\n".join(lines[-25:]) if lines else ""
                            if len(tail) > 2000:
                                tail = tail[-2000:]
                            detail = tail or f"Код выхода: {code}"
                        except OSError:
                            detail = f"Код выхода: {code}"
                except Exception as exc:
                    detail = str(exc)

                code = proc.returncode if proc is not None else -1
                if code == 2:
                    ok = False
                    msg = "Отчёты не сформированы: нет входных данных (см. лог и пути к нулевым файлам)"
                elif ok:
                    msg = "Готово"
                else:
                    msg = "Формирование отчёта завершилось с ошибкой"

                with open(STATUS_FILE, "w", encoding="utf-8") as f:
                    json.dump(
                        {
                            "done": True,
                            "ok": ok,
                            "percent": 100,
                            "message": msg,
                            "detail": detail,
                        },
                        f,
                        ensure_ascii=False,
                    )
                try:
                    BUILD_PROGRESS_FILE.unlink(missing_ok=True)
                    BUILD_PROGRESS_FILE.with_suffix(".json.tmp").unlink(missing_ok=True)
                except OSError:
                    pass
        
        for stale in (BUILD_PROGRESS_FILE, BUILD_PROGRESS_FILE.with_suffix(".json.tmp")):
            try:
                stale.unlink(missing_ok=True)
            except OSError:
                pass

        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "done": False,
                    "ok": None,
                    "percent": 0,
                    "message": "Запуск…",
                    "detail": "",
                    "started_at": time.time(),
                },
                f,
                ensure_ascii=False,
            )

        threading.Thread(target=run, daemon=True).start()
        return "OK", 200

    @app.route("/status")
    def get_status():
        default_running = {
            "done": False,
            "ok": None,
            "percent": 0,
            "message": "Запуск…",
            "detail": "",
            "started_at": None,
        }
        try:
            with open(STATUS_FILE, "r", encoding="utf-8") as f:
                st = json.load(f)
        except (OSError, json.JSONDecodeError, UnicodeDecodeError, ValueError, TypeError):
            st = dict(default_running)

        if st.get("done"):
            out = {k: v for k, v in st.items() if k != "started_at"}
            return out

        started = st.get("started_at")
        elapsed = 0
        if isinstance(started, (int, float)):
            elapsed = max(0, int(time.time() - float(started)))

        pct = int(st.get("percent", 0) or 0)
        msg = st.get("message") or "Запуск…"
        try:
            if BUILD_PROGRESS_FILE.is_file():
                with open(BUILD_PROGRESS_FILE, "r", encoding="utf-8") as pf:
                    prog = json.load(pf)
                pct = int(prog.get("percent", pct))
                msg = prog.get("message") or msg
        except (OSError, json.JSONDecodeError, UnicodeDecodeError, ValueError, TypeError):
            pass

        tail = ""
        try:
            if BUILD_LOG.is_file():
                txt = BUILD_LOG.read_text(encoding="utf-8", errors="replace")
                lines = [ln for ln in txt.splitlines() if ln.strip()]
                if lines:
                    tail = lines[-1][-500:]
        except OSError:
            pass

        detail_lines = [f"Прошло {elapsed} с"]
        if tail:
            detail_lines.append(tail)
        return {
            "done": False,
            "ok": None,
            "percent": max(0, min(99, pct)),
            "message": msg,
            "detail": "\n".join(detail_lines),
        }

    @app.route("/save_exceptions", methods=["POST"])
    def save_exceptions():
        try:
            import pandas as pd
            data = request.get_json() or {}
            exceptions = data.get("exceptions", [])
            
            if not exceptions:
                return jsonify({"error": "Нет исключений"}), 400
            existing = _load_exceptions_df()
            
            new_data = []
            for exc in exceptions:
                so = _norm(exc.get("so", ""))
                customer = _norm(exc.get("customer", ""))
                if not so or not customer:
                    continue
                new_data.append({
                    "SO": so,
                    "Customer": customer,
                    "Comment OM": "Добавлено через веб-интерфейс"
                })
            if not new_data:
                return jsonify({"error": "Нет валидных SO/Customer"}), 400
            
            combined = pd.concat([existing, pd.DataFrame(new_data)], ignore_index=True)
            combined["SO"] = combined["SO"].map(_norm)
            combined["Customer"] = combined["Customer"].map(_norm)
            exc_file = exception_file_path()
            exc_file.parent.mkdir(parents=True, exist_ok=True)
            combined.to_excel(exc_file, index=False)
            
            return jsonify({"success": True, "saved": len(new_data), "total": int(len(combined))})
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    @app.route("/exceptions", methods=["GET"])
    def list_exceptions():
        try:
            df = _load_exceptions_df()
            df["SO"] = df["SO"].map(_norm)
            df["Customer"] = df["Customer"].map(_norm)
            rows = (
                df[df["SO"].astype(str).str.strip().ne("") & df["Customer"].astype(str).str.strip().ne("")]
                .drop_duplicates(subset=["SO", "Customer"], keep="last")
            )
            out = [{"so": str(r["SO"]), "customer": str(r["Customer"])} for _, r in rows.iterrows()]
            return jsonify({"success": True, "exceptions": out, "total": len(out)})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/exceptions/delete", methods=["POST"])
    def delete_exceptions():
        try:
            data = request.get_json() or {}
            exceptions = data.get("exceptions", [])
            if not exceptions:
                return jsonify({"success": False, "error": "Нет исключений для удаления"}), 400
            targets = {(_norm(x.get("so", "")), _norm(x.get("customer", ""))) for x in exceptions}
            targets = {t for t in targets if t[0] and t[1]}
            if not targets:
                return jsonify({"success": False, "error": "Нет валидных SO/Customer"}), 400

            df = _load_exceptions_df()
            so_s = df["SO"].map(_norm)
            cust_s = df["Customer"].map(_norm)
            mask_drop = [(a, b) in targets for a, b in zip(so_s, cust_s)]
            remaining = df[[not x for x in mask_drop]].copy()
            exc_file = exception_file_path()
            exc_file.parent.mkdir(parents=True, exist_ok=True)
            remaining.to_excel(exc_file, index=False)
            return jsonify({"success": True, "deleted": int(sum(mask_drop)), "total": int(len(remaining))})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/paths", methods=["POST"])
    def save_paths():
        try:
            data = request.get_json() or {}
            base_dir = _normalize_path_text(str(data.get("base_dir", "")))
            output_dir = _normalize_path_text(str(data.get("output_dir", "")))
            exception_file = _normalize_path_text(str(data.get("exception_file", "")))
            if not base_dir or not output_dir or not exception_file:
                return jsonify({"success": False, "error": "Все пути обязательны"}), 400
            base_p = Path(base_dir)
            out_p = Path(output_dir)
            exc_p = Path(exception_file)
            if not base_p.exists():
                return jsonify({"success": False, "error": f"Каталог нулевых выгрузок не найден: {base_p}"}), 400
            if not base_p.is_dir():
                return jsonify({"success": False, "error": f"Путь нулевых выгрузок не является каталогом: {base_p}"}), 400
            out_p.mkdir(parents=True, exist_ok=True)
            exc_p.parent.mkdir(parents=True, exist_ok=True)
            payload = _save_runtime_paths(base_dir, output_dir, exception_file)
            return jsonify({"success": True, **payload})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/favicon.ico")
    def favicon():
        return Response(status=204)

    return app


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    host = "127.0.0.1"
    port = 8765
    url = f"http://{host}:{port}/"

    _probe_port_available(host, port)

    if not args.no_browser:
        threading.Thread(target=lambda: (time.sleep(1), webbrowser.open(url)), daemon=True).start()

    print(f"Откройте в браузере: {url}")

    app = create_app()
    app.run(host=host, port=port, debug=False, use_reloader=False, threaded=True)


if __name__ == "__main__":
    main()


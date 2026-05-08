# -*- coding: utf-8 -*-
"""
Служба Windows: веб проверки PF BP-PY-ZY (Waitress), по умолчанию http://127.0.0.1:8765/

От администратора:
  python -m pip install pywin32
  python pf_checks_windows_service.py install
  python pf_checks_windows_service.py --startup auto update
  net start PFBPYYZYWeb

Или запустите install_windows_service.cmd
"""
from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path

SERVICE_ROOT = Path(__file__).resolve().parent

try:
    import win32event
    import win32service
    import win32serviceutil
    import servicemanager
except ImportError as exc:
    raise SystemExit("Нужен pywin32: python -m pip install pywin32") from exc


class PfBpPyZyWebService(win32serviceutil.ServiceFramework):
    _svc_name_ = "PFBPYYZYWeb"
    _svc_display_name_ = "PF BP-PY-ZY — веб-проверка (localhost)"
    _svc_description_ = (
        "Веб-интерфейс проверки PF BP-PY-ZY на этой машине (http://127.0.0.1:8765/). "
        "Папка с данными — каталог установки службы (build_checks, нулевые выгрузки)."
    )

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, ""),
        )
        os.environ.setdefault("REPORTS_WEB_SERVER", "1")
        os.environ.setdefault("REPORTS_WEB_HOST", "127.0.0.1")
        os.environ.setdefault("REPORTS_WEB_NO_BROWSER", "1")
        os.chdir(str(SERVICE_ROOT))
        if str(SERVICE_ROOT) not in sys.path:
            sys.path.insert(0, str(SERVICE_ROOT))

        try:
            from run_checks_web import create_app
            try:
                from run_checks_web import _append_startup_log  # type: ignore
            except Exception:
                def _append_startup_log(message: str) -> None:  # type: ignore
                    try:
                        with open(SERVICE_ROOT / "run_checks_web_startup.log", "a", encoding="utf-8") as f:
                            f.write(message + "\n")
                    except OSError:
                        pass
            from waitress import serve

            _append_startup_log("Windows Service: старт Waitress")

            app = create_app()
            host = os.environ.get("REPORTS_WEB_HOST", "127.0.0.1")
            port = int(os.environ.get("REPORTS_WEB_PORT", "8765"))

            def runner() -> None:
                try:
                    serve(app, host=host, port=port, threads=4, channel_timeout=7200)
                except Exception as run_exc:
                    try:
                        _append_startup_log(f"Waitress: {run_exc!r}")
                    except OSError:
                        pass

            threading.Thread(target=runner, daemon=True, name="Waitress").start()
        except Exception:
            err = traceback.format_exc()
            try:
                (SERVICE_ROOT / "run_checks_web_error.log").write_text(err, encoding="utf-8")
            except OSError:
                pass
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_ERROR_TYPE,
                servicemanager.PYS_SERVICE_STOPPED,
                (self._svc_name_, "Ошибка запуска — см. run_checks_web_error.log в каталоге службы."),
            )
            return

        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(PfBpPyZyWebService)

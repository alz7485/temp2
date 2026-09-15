import sys, traceback
from pathlib import Path
from PySide6.QtWidgets import QApplication, QMessageBox
from main_window import MainWindow
from settings_manager import SettingsManager
from logger import setup_logger, log_exception

def get_base_dir():
    return Path(sys.executable if getattr(sys,"frozen",False) else __file__).resolve().parent

def main():
    app=QApplication(sys.argv); app.setApplicationName("Outlook DB Builder")
    base=get_base_dir(); setup_logger(base)
    w=MainWindow(SettingsManager(base)); w.show(); return app.exec()

if __name__ == "__main__":
    try: sys.exit(main())
    except Exception:
        err=traceback.format_exc()
        try: (get_base_dir()/"error.log").write_text(err, encoding="utf-8")
        except Exception: pass
        log_exception("Fatal application error")
        app=QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(None,"Outlook DB Builder","起動中にエラーが発生しました。\n\n"+err[-1500:])
        sys.exit(1)

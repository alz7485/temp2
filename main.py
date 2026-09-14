import sys
import traceback
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMessageBox

from main_window import MainWindow
from settings_manager import SettingsManager


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def write_error_log(error_text: str):
    try:
        log_path = get_base_dir() / "error.log"
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(error_text)
    except Exception:
        pass


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Outlook DB Builder")

    base_dir = get_base_dir()
    settings_manager = SettingsManager(base_dir)

    window = MainWindow(settings_manager)
    window.show()

    return app.exec()


if __name__ == "__main__":
    try:
        exit_code = main()

    except Exception:
        error_text = traceback.format_exc()
        write_error_log(error_text)

        try:
            print(error_text)
        except Exception:
            pass

        try:
            app = QApplication.instance()
            if app is None:
                app = QApplication(sys.argv)

            QMessageBox.critical(
                None,
                "Outlook DB Builder",
                "起動中にエラーが発生しました。\n\n"
                "exeと同じフォルダの error.log を確認してください。\n\n"
                + error_text[-1500:],
            )
        except Exception:
            pass

        sys.exit(1)

    sys.exit(exit_code)

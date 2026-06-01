"""
Точка входа. Запускает GUI или CLI в зависимости от аргументов.

  python main.py                          → GUI
  python main.py check.docx              → CLI, консольный отчёт
  python main.py check.docx -o rep.html  → CLI, HTML-отчёт
"""
import sys


def main():
    # Если передан аргумент-файл — CLI режим
    if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
        from vkr_checker.cli import main as cli_main
        cli_main()
    else:
        try:
            from vkr_checker.gui import run_gui
            run_gui()
        except ImportError:
            # tkinter недоступен — падаем в CLI
            from vkr_checker.cli import main as cli_main
            cli_main()


if __name__ == "__main__":
    main()

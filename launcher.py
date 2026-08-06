from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
VENV_DIR = BASE_DIR / ".venv"
VENV_PYTHON = VENV_DIR / "Scripts" / "python.exe"
REQUIREMENTS = BASE_DIR / "requirements.txt"
APP = BASE_DIR / "app.py"


def notify_error(message: str) -> None:
    try:
        from tkinter import messagebox

        messagebox.showerror("학생부 기록 정리", message)
    except Exception:
        pass
    print(f"\n[오류] {message}")


def run_checked(command: list[str]) -> None:
    subprocess.run(command, cwd=BASE_DIR, check=True)


def dependencies_ready() -> bool:
    if not VENV_PYTHON.is_file():
        return False
    result = subprocess.run(
        [str(VENV_PYTHON), "-c", "import flask, openpyxl"],
        cwd=BASE_DIR,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    return result.returncode == 0


def main() -> int:
    try:
        if not VENV_PYTHON.is_file():
            print("처음 실행을 준비하고 있습니다. 잠시 기다려 주세요...")
            run_checked([sys.executable, "-m", "venv", str(VENV_DIR)])

        if not dependencies_ready():
            print("필요한 구성 요소를 설치하고 있습니다...")
            run_checked(
                [
                    str(VENV_PYTHON),
                    "-m",
                    "pip",
                    "install",
                    "--disable-pip-version-check",
                    "-r",
                    str(REQUIREMENTS),
                ]
            )

        print("\n웹앱을 시작합니다: http://127.0.0.1:8765")
        print("종료하려면 이 창에서 Ctrl+C를 누르세요.\n")
        environment = os.environ.copy()
        environment["OPEN_BROWSER"] = "0" if os.environ.get("NO_BROWSER") == "1" else "1"
        return subprocess.call(
            [str(VENV_PYTHON), str(APP)],
            cwd=BASE_DIR,
            env=environment,
        )
    except KeyboardInterrupt:
        return 0
    except subprocess.CalledProcessError as exc:
        notify_error(
            "실행 준비 중 문제가 발생했습니다. 인터넷 연결을 확인한 뒤 다시 실행해 주세요.\n"
            f"오류 코드: {exc.returncode}"
        )
        return 1
    except Exception as exc:
        notify_error(f"웹앱을 실행하지 못했습니다.\n{exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

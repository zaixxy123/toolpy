import argparse
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path


def _parse_args():
    parser = argparse.ArgumentParser(
        add_help=False
    )
    parser.add_argument(
        "--toolpy-updater",
        action="store_true",
    )
    parser.add_argument(
        "--pid",
        type=int,
        required=True,
    )
    parser.add_argument(
        "--current",
        required=True,
    )
    parser.add_argument(
        "--update",
        required=True,
    )
    parser.add_argument(
        "--helper",
        required=True,
    )
    parser.add_argument(
        "--log",
        required=True,
    )
    return parser.parse_args()


def _process_exists(pid):
    try:
        result = subprocess.run(
            [
                "tasklist",
                "/FI",
                f"PID eq {pid}",
                "/NH",
            ],
            capture_output=True,
            text=True,
            creationflags=getattr(
                subprocess,
                "CREATE_NO_WINDOW",
                0,
            ),
            timeout=5,
        )
        return str(pid) in result.stdout
    except Exception:
        return True


def _validate_exe(path):
    try:
        if (
            not path.exists()
            or path.stat().st_size < 1024 * 1024
        ):
            return False

        with path.open("rb") as file:
            return file.read(2) == b"MZ"

    except OSError:
        return False


def _show_error(log_path, message):
    try:
        log_path.write_text(
            message,
            encoding="utf-8",
        )
        subprocess.Popen(
            ["notepad.exe", str(log_path)]
        )
    except Exception:
        pass


def _schedule_helper_cleanup(helper_path):
    cleanup_batch = (
        Path(tempfile.gettempdir())
        / "ToolPy_helper_cleanup.bat"
    )

    cleanup_batch.write_text(
        "@echo off\n"
        "timeout /t 2 /nobreak >nul\n"
        f'del /F /Q "{helper_path}" >nul 2>&1\n'
        'del /F /Q "%~f0" >nul 2>&1\n',
        encoding="utf-8",
    )

    subprocess.Popen(
        [
            "cmd.exe",
            "/c",
            "start",
            "",
            "/min",
            str(cleanup_batch),
        ],
        creationflags=getattr(
            subprocess,
            "CREATE_NO_WINDOW",
            0,
        ),
        close_fds=True,
    )


def run_updater_mode():
    args = _parse_args()

    current = Path(
        args.current
    ).resolve()
    update = Path(
        args.update
    ).resolve()
    helper = Path(
        args.helper
    ).resolve()
    log = Path(
        args.log
    ).resolve()

    backup = current.with_suffix(
        current.suffix + ".old"
    )

    try:
        if not _validate_exe(update):
            raise RuntimeError(
                "The downloaded update is not a valid EXE."
            )

        for _ in range(120):
            if not _process_exists(args.pid):
                break

            time.sleep(0.5)
        else:
            raise RuntimeError(
                "ToolPy did not close in time."
            )

        time.sleep(2)

        installed = False
        last_error = None

        for _ in range(45):
            try:
                backup.unlink(
                    missing_ok=True
                )

                if current.exists():
                    os.replace(
                        current,
                        backup,
                    )

                os.replace(
                    update,
                    current,
                )

                if not _validate_exe(current):
                    raise RuntimeError(
                        "The replacement EXE failed validation."
                    )

                installed = True
                break

            except Exception as error:
                last_error = error

                try:
                    current.unlink(
                        missing_ok=True
                    )
                except OSError:
                    pass

                try:
                    if backup.exists():
                        os.replace(
                            backup,
                            current,
                        )
                except OSError:
                    pass

                time.sleep(1)

        if not installed:
            raise RuntimeError(
                "ToolPy could not replace the old EXE."
                + (
                    f"\n\nLast error: {last_error}"
                    if last_error is not None
                    else ""
                )
            )

        subprocess.Popen(
            [str(current)],
            cwd=str(current.parent),
            close_fds=True,
        )

        time.sleep(3)

        try:
            backup.unlink(
                missing_ok=True
            )
        except OSError:
            pass

        try:
            log.unlink(
                missing_ok=True
            )
        except OSError:
            pass

        _schedule_helper_cleanup(helper)
        return 0

    except Exception as error:
        try:
            if (
                not current.exists()
                and backup.exists()
            ):
                os.replace(
                    backup,
                    current,
                )
        except OSError:
            pass

        _show_error(
            log,
            "ToolPy update failed.\n\n"
            f"Error: {error}\n\n"
            f"Current EXE: {current}\n"
            f"Downloaded EXE: {update}\n",
        )

        _schedule_helper_cleanup(helper)
        return 1

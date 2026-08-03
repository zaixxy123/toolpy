import json
import os
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import uuid
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QMessageBox,
    QProgressDialog,
)

from utils import resource_path


GITHUB_OWNER = "zaixxy123"
GITHUB_REPO = "toolpy"
RELEASE_API = (
    f"https://api.github.com/repos/"
    f"{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
)
EXE_ASSET_NAME = "ToolPy.exe"
USER_AGENT = "ToolPy-AutoUpdater/1.0"


def _version_tuple(value):
    clean = value.strip().lower().lstrip("v")
    parts = []

    for piece in clean.split("."):
        digits = "".join(
            character
            for character in piece
            if character.isdigit()
        )
        parts.append(int(digits or 0))

    while len(parts) < 3:
        parts.append(0)

    return tuple(parts[:3])


def _current_version():
    try:
        version_file = Path(resource_path("version.txt"))
        return version_file.read_text(
            encoding="utf-8"
        ).strip()
    except Exception:
        return "0.0.0"


def _is_frozen_exe():
    return bool(getattr(sys, "frozen", False))


def _powershell_quote(value):
    return str(value).replace("'", "''")


class UpdateCheckWorker(QThread):
    update_available = Signal(str, str)
    no_update = Signal()
    failed = Signal(str)

    def run(self):
        try:
            request = urllib.request.Request(
                RELEASE_API,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": USER_AGENT,
                    "X-GitHub-Api-Version": "2022-11-28",
                    "Cache-Control": "no-cache",
                },
            )

            with urllib.request.urlopen(
                request,
                timeout=15,
            ) as response:
                release = json.load(response)

            latest_version = str(
                release.get("tag_name", "")
            ).strip()

            download_url = ""

            for asset in release.get("assets", []):
                if asset.get("name") == EXE_ASSET_NAME:
                    download_url = str(
                        asset.get(
                            "browser_download_url",
                            "",
                        )
                    ).strip()
                    break

            if not latest_version:
                raise RuntimeError(
                    "The latest GitHub release has no version tag."
                )

            if not download_url:
                raise RuntimeError(
                    "The latest release does not contain "
                    f"{EXE_ASSET_NAME}."
                )

            if (
                _version_tuple(latest_version)
                > _version_tuple(_current_version())
            ):
                self.update_available.emit(
                    latest_version,
                    download_url,
                )
            else:
                self.no_update.emit()

        except urllib.error.HTTPError as error:
            self.failed.emit(
                f"GitHub returned HTTP {error.code}."
            )
        except urllib.error.URLError as error:
            reason = getattr(error, "reason", "")
            self.failed.emit(
                "Could not connect to GitHub."
                + (f"\n{reason}" if reason else "")
            )
        except Exception as error:
            self.failed.emit(str(error))


class DownloadWorker(QThread):
    progress = Signal(int)
    completed = Signal(str)
    failed = Signal(str)

    def __init__(self, download_url):
        super().__init__()
        self.download_url = download_url

    def run(self):
        destination = None

        try:
            update_dir = (
                Path(tempfile.gettempdir())
                / "ToolPyUpdater"
            )
            update_dir.mkdir(
                parents=True,
                exist_ok=True,
            )

            destination = (
                update_dir
                / f"ToolPy_update_{uuid.uuid4().hex}.exe"
            )

            request = urllib.request.Request(
                self.download_url,
                headers={
                    "User-Agent": USER_AGENT,
                    "Accept": "application/octet-stream",
                    "Cache-Control": "no-cache",
                },
            )

            with urllib.request.urlopen(
                request,
                timeout=120,
            ) as response:
                total_size = int(
                    response.headers.get(
                        "Content-Length",
                        "0",
                    )
                    or 0
                )
                downloaded = 0

                with destination.open("wb") as output:
                    while True:
                        chunk = response.read(
                            1024 * 1024
                        )

                        if not chunk:
                            break

                        output.write(chunk)
                        downloaded += len(chunk)

                        if total_size > 0:
                            percentage = int(
                                downloaded
                                * 100
                                / total_size
                            )
                            self.progress.emit(
                                min(percentage, 99)
                            )

            if not destination.exists():
                raise RuntimeError(
                    "The downloaded update file was not created."
                )

            file_size = destination.stat().st_size

            if file_size < 1024 * 1024:
                raise RuntimeError(
                    "The downloaded update is unexpectedly small."
                )

            with destination.open("rb") as downloaded_file:
                if downloaded_file.read(2) != b"MZ":
                    raise RuntimeError(
                        "The downloaded file is not a valid Windows EXE."
                    )

            self.progress.emit(100)
            self.completed.emit(str(destination))

        except Exception as error:
            if destination is not None:
                try:
                    destination.unlink(missing_ok=True)
                except Exception:
                    pass

            self.failed.emit(str(error))


class UpdateManager(QObject):
    def __init__(self):
        super().__init__()

        self.parent = None
        self.check_worker = None
        self.download_worker = None
        self.progress_dialog = None

    def check_for_updates(self, parent):
        self.parent = parent

        if not _is_frozen_exe():
            return

        if (
            self.check_worker is not None
            and self.check_worker.isRunning()
        ):
            return

        self.check_worker = UpdateCheckWorker()
        self.check_worker.update_available.connect(
            self._offer_update
        )
        self.check_worker.failed.connect(
            self._check_failed
        )
        self.check_worker.finished.connect(
            self._clear_check_worker
        )
        self.check_worker.start()

    def _clear_check_worker(self):
        self.check_worker = None

    def _clear_download_worker(self):
        self.download_worker = None

    def _offer_update(
        self,
        latest_version,
        download_url,
    ):
        answer = QMessageBox.question(
            self.parent,
            "Update Available",
            "A new ToolPy version is available.\n\n"
            f"Current: v{_current_version().lstrip('v')}\n"
            f"Latest:  {latest_version}\n\n"
            "Download and install it now?",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )

        if answer == QMessageBox.StandardButton.Yes:
            self._download_update(download_url)

    def _download_update(self, download_url):
        self.progress_dialog = QProgressDialog(
            "Downloading update...",
            "",
            0,
            100,
            self.parent,
        )
        self.progress_dialog.setWindowTitle(
            "Updating ToolPy"
        )
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setAutoReset(False)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.setValue(0)
        self.progress_dialog.show()

        self.download_worker = DownloadWorker(
            download_url
        )
        self.download_worker.progress.connect(
            self.progress_dialog.setValue
        )
        self.download_worker.completed.connect(
            self._install_and_restart
        )
        self.download_worker.failed.connect(
            self._download_failed
        )
        self.download_worker.finished.connect(
            self._clear_download_worker
        )
        self.download_worker.start()

    def _install_and_restart(
        self,
        downloaded_exe,
    ):
        if self.progress_dialog:
            self.progress_dialog.setLabelText(
                "Installing and restarting..."
            )
            self.progress_dialog.setValue(100)

        try:
            current_exe = Path(
                sys.executable
            ).resolve()
            downloaded_exe = Path(
                downloaded_exe
            ).resolve()

            if not current_exe.exists():
                raise RuntimeError(
                    "ToolPy could not locate its current EXE."
                )

            if not downloaded_exe.exists():
                raise RuntimeError(
                    "The downloaded update file is missing."
                )

            update_dir = downloaded_exe.parent
            script_file = (
                update_dir
                / f"install_{uuid.uuid4().hex}.ps1"
            )
            log_file = (
                update_dir
                / "ToolPy_update_error.txt"
            )
            backup_exe = current_exe.with_suffix(
                current_exe.suffix + ".old"
            )

            current_pid = os.getpid()
            working_directory = current_exe.parent

            current_ps = _powershell_quote(current_exe)
            update_ps = _powershell_quote(downloaded_exe)
            backup_ps = _powershell_quote(backup_exe)
            log_ps = _powershell_quote(log_file)
            workdir_ps = _powershell_quote(
                working_directory
            )

            script = f"""$ErrorActionPreference = 'Stop'

$current = '{current_ps}'
$update = '{update_ps}'
$backup = '{backup_ps}'
$log = '{log_ps}'
$workingDirectory = '{workdir_ps}'
$processId = {current_pid}

function Restore-OldVersion {{
    try {{
        if (-not (Test-Path -LiteralPath $current) -and
            (Test-Path -LiteralPath $backup)) {{
            Move-Item -LiteralPath $backup `
                -Destination $current -Force
        }}
    }} catch {{
    }}
}}

try {{
    for ($i = 0; $i -lt 120; $i++) {{
        $process = Get-Process -Id $processId `
            -ErrorAction SilentlyContinue

        if ($null -eq $process) {{
            break
        }}

        Start-Sleep -Milliseconds 500
    }}

    if (Get-Process -Id $processId `
        -ErrorAction SilentlyContinue) {{
        throw 'ToolPy did not close in time.'
    }}

    $installed = $false

    for ($attempt = 1; $attempt -le 30; $attempt++) {{
        try {{
            if (Test-Path -LiteralPath $backup) {{
                Remove-Item -LiteralPath $backup `
                    -Force -ErrorAction SilentlyContinue
            }}

            if (Test-Path -LiteralPath $current) {{
                Move-Item -LiteralPath $current `
                    -Destination $backup -Force
            }}

            Copy-Item -LiteralPath $update `
                -Destination $current -Force

            if (-not (Test-Path -LiteralPath $current)) {{
                throw 'The new ToolPy.exe was not created.'
            }}

            $newSize = (
                Get-Item -LiteralPath $current
            ).Length

            if ($newSize -lt 1048576) {{
                throw 'The new ToolPy.exe is unexpectedly small.'
            }}

            $installed = $true
            break
        }} catch {{
            Restore-OldVersion
            Start-Sleep -Seconds 1
        }}
    }}

    if (-not $installed) {{
        throw 'ToolPy could not replace the old EXE.'
    }}

    Start-Process -FilePath $current `
        -WorkingDirectory $workingDirectory

    Start-Sleep -Seconds 2

    Remove-Item -LiteralPath $backup `
        -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $update `
        -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $log `
        -Force -ErrorAction SilentlyContinue
}} catch {{
    Restore-OldVersion

    $message = @(
        'ToolPy update failed.'
        ''
        ('Error: ' + $_.Exception.Message)
        ('Current EXE: ' + $current)
        ('Downloaded EXE: ' + $update)
    ) -join [Environment]::NewLine

    Set-Content -LiteralPath $log `
        -Value $message -Encoding UTF8

    Start-Process -FilePath 'notepad.exe' `
        -ArgumentList $log
}} finally {{
    Start-Sleep -Seconds 1
    Remove-Item -LiteralPath $PSCommandPath `
        -Force -ErrorAction SilentlyContinue
}}
"""

            script_file.write_text(
                script,
                encoding="utf-8-sig",
            )

            powershell = (
                Path(
                    os.environ.get(
                        "SystemRoot",
                        r"C:\Windows",
                    )
                )
                / "System32"
                / "WindowsPowerShell"
                / "v1.0"
                / "powershell.exe"
            )

            if not powershell.exists():
                raise RuntimeError(
                    "Windows PowerShell was not found."
                )

            creation_flags = (
                getattr(
                    subprocess,
                    "CREATE_NO_WINDOW",
                    0,
                )
                | getattr(
                    subprocess,
                    "DETACHED_PROCESS",
                    0,
                )
            )

            subprocess.Popen(
                [
                    str(powershell),
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-WindowStyle",
                    "Hidden",
                    "-File",
                    str(script_file),
                ],
                creationflags=creation_flags,
                close_fds=True,
                cwd=str(working_directory),
            )

            QApplication.quit()

        except Exception as error:
            if self.progress_dialog:
                self.progress_dialog.close()

            QMessageBox.critical(
                self.parent,
                "Update Failed",
                "ToolPy downloaded the update but "
                "could not install it.\n\n"
                f"{error}",
            )

    def _check_failed(self, message):
        print(
            f"Update check skipped: {message}"
        )

    def _download_failed(self, message):
        if self.progress_dialog:
            self.progress_dialog.close()

        QMessageBox.critical(
            self.parent,
            "Download Failed",
            "ToolPy could not download the update.\n\n"
            f"{message}",
        )

# 在 WSL 中把脚本转交给 Windows 侧 Python 执行，Visio COM 只能由 Windows 进程调用。

import glob
import os
import platform
import subprocess
import sys


VENV_GLOB = "/mnt/c/Users/*/.venvs/visio-skill/Scripts/python.exe"
FALLBACK_GLOBS = [
    "/mnt/c/Program Files/Python3*/python.exe",
    "/mnt/c/Users/*/AppData/Local/Programs/Python/Python3*/python.exe",
]


def running_in_wsl() -> bool:
    if sys.platform != "linux":
        return False
    return "microsoft" in platform.release().lower()


def _has_win32com(python_exe: str) -> bool:
    result = subprocess.run(
        [python_exe, "-c", "import win32com.client"],
        capture_output=True,
    )
    return result.returncode == 0


def _candidates() -> list[str]:
    found = []
    override = os.environ.get("VISIO_PYTHON")
    if override:
        found.append(override)
    found.extend(sorted(glob.glob(VENV_GLOB)))
    for pattern in FALLBACK_GLOBS:
        found.extend(sorted(glob.glob(pattern)))
    return found


def find_windows_python() -> str:
    tried = []
    for candidate in _candidates():
        if not os.path.exists(candidate):
            continue
        tried.append(candidate)
        if _has_win32com(candidate):
            return candidate
    detail = "\n".join(f"  - {path}" for path in tried) or "  （没有找到任何 python.exe）"
    raise RuntimeError(
        "找不到装有 pywin32 的 Windows Python，无法调用 Visio COM。\n"
        f"已尝试的解释器：\n{detail}\n"
        "安装方法：python.exe -m venv %USERPROFILE%\\.venvs\\visio-skill，"
        "再对该环境执行 pip install pywin32；或用环境变量 VISIO_PYTHON 指定解释器。"
    )


def to_windows_path(path: str) -> str:
    result = subprocess.run(["wslpath", "-w", path], capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"无法转换为 Windows 路径：{path}\n{result.stderr.strip()}")
    return result.stdout.strip()


def _convert_arg(arg: str) -> str:
    if arg.startswith("--") and "=" in arg:
        flag, value = arg.split("=", 1)
        if value.startswith("/"):
            return f"{flag}={to_windows_path(value)}"
        return arg
    if arg.startswith("/") and not arg.startswith("//"):
        return to_windows_path(arg)
    return arg


def ensure_windows_runtime(entry_script: str) -> None:
    if not running_in_wsl():
        return
    python_exe = find_windows_python()
    argv = [python_exe, "-X", "utf8", to_windows_path(os.path.abspath(entry_script))]
    argv.extend(_convert_arg(arg) for arg in sys.argv[1:])
    os.execve(python_exe, argv, os.environ.copy())

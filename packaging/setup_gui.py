# -*- coding: utf-8 -*-
"""
AttachLink 安装器（GUI 外壳）
- 业务逻辑全部走内嵌的 PowerShell 脚本（setup.ps1 / install-elevated.ps1），
  本程序只负责界面、日志展示与提权桥接，不复制业务逻辑（单一事实源）。
- 静默参数（供自动化验证/高级用户）：
    AttachLinkSetup.exe /install             标准安装（用户级）
    AttachLinkSetup.exe /deploy-elevated     鼨信诺管理员环境部署（弹 UAC）
    AttachLinkSetup.exe /uninstall           卸载用户级
    AttachLinkSetup.exe /uninstall-elevated  卸载机器级（弹 UAC）
    AttachLinkSetup.exe /selfcheck           自检（校验内嵌文件完整性）
  退出码：0 成功；2 Excel 在运行；3 安装包不完整；1 其他失败。
"""
import os
import sys
import queue
import threading
import subprocess
import tkinter as tk
from tkinter import scrolledtext
from tkinter import ttk

VERSION_FILE = os.path.join("payload", "VERSION.txt")
APP_TITLE = "AttachLink 安装器"


def payload_dir() -> str:
    """PyInstaller onefile 运行时解包目录；源码直跑时用 packaging/payload"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, "payload")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "payload")


def read_version() -> str:
    try:
        with open(os.path.join(payload_dir(), "VERSION.txt"), encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return "未知版本"


def log_file_path() -> str:
    return os.path.join(os.environ.get("TEMP", "."), "AttachLinkSetup.log")


def append_log_file(text: str) -> None:
    try:
        with open(log_file_path(), "a", encoding="utf-8") as f:
            f.write(text + "\n")
    except OSError:
        pass


class Runner:
    """在后台线程运行 PowerShell，把 stdout 逐行送入队列"""

    def __init__(self, outq: "queue.Queue[str]") -> None:
        self.outq = outq
        self.proc: subprocess.Popen | None = None

    def run(self, args: list[str], cwd: str) -> int:
        self.outq.put("$ " + " ".join(args[:4]) + " ...\n")
        self.proc = subprocess.Popen(
            args,
            cwd=cwd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        assert self.proc.stdout is not None
        for line in self.proc.stdout:
            self.outq.put(line)
        self.proc.stdout.close()
        return self.proc.wait()

    def run_elevated(self, helper: str, target: str) -> int:
        """经 run-elevated.ps1 提权运行目标脚本：UAC 由系统弹出，日志回传临时文件"""
        return self.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", helper, target],
            payload_dir(),
        )


def ps(script_rel: str, extra: list[str] | None = None) -> list[str]:
    args = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
            os.path.join(payload_dir(), script_rel)]
    return args + (extra or [])


class App:
    def __init__(self) -> None:
        self.ver = read_version()
        self.outq: "queue.Queue[str]" = queue.Queue()
        self.runner: Runner | None = None
        self.worker: threading.Thread | None = None
        self.root = tk.Tk()
        self.root.title(f"{APP_TITLE} v{self.ver}")
        self.root.geometry("760x520")
        self.root.minsize(640, 420)
        self._build_ui()
        self.root.after(100, self._drain_queue)
        self.log(f"AttachLink v{self.ver}　|　安装日志：{log_file_path()}")

    # ---------------- UI ----------------
    def _build_ui(self) -> None:
        head = ttk.Frame(self.root, padding=(12, 10, 12, 0))
        head.pack(fill="x")
        ttk.Label(head, text=f"AttachLink 审计底稿附件链接加载项　v{self.ver}",
                  font=("Microsoft YaHei UI", 13, "bold")).pack(anchor="w")
        ttk.Label(head, text="接收/复核归档底稿的同事：只需「标准安装」。",
                  foreground="#444").pack(anchor="w")
        ttk.Label(head, text="在本机以管理员运行鼎信诺编制底稿的同事：标准安装后，再执行「鼎信诺管理员环境部署」。",
                  foreground="#444").pack(anchor="w")

        bar = ttk.Frame(self.root, padding=(12, 8, 12, 4))
        bar.pack(fill="x")
        self.buttons: dict[str, ttk.Button] = {}
        for key, text, cmd in [
            ("install", "标准安装（接收 / 归档同事）", self.act_install),
            ("elev", "鼎信诺管理员环境部署", self.act_elevated),
            ("uninst", "卸载（用户级）", self.act_uninstall),
            ("unelev", "卸载（机器级）", self.act_uninstall_elevated),
        ]:
            b = ttk.Button(bar, text=text, command=cmd)
            b.pack(side="left", padx=(0, 8))
            self.buttons[key] = b

        self.logbox = scrolledtext.ScrolledText(self.root, wrap="word", state="disabled",
                                                font=("Consolas", 10), background="#fafafa")
        self.logbox.pack(fill="both", expand=True, padx=12, pady=(4, 4))
        self.status = ttk.Label(self.root, text="就绪", padding=(12, 0, 12, 8), relief="sunken", anchor="w")
        self.status.pack(fill="x")

    def log(self, text: str) -> None:
        append_log_file(text.rstrip("\n"))
        self.logbox.configure(state="normal")
        self.logbox.insert("end", text if text.endswith("\n") else text + "\n")
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    # ---------------- actions ----------------
    def _start(self, fn, *a) -> None:
        if self.worker and self.worker.is_alive():
            return
        for b in self.buttons.values():
            b.configure(state="disabled")
        self.status.configure(text="执行中…")

        def wrap() -> None:
            try:
                code = fn(*a)
                self.outq.put(f"\n[退出码 {code}] " + ("操作成功。" if code == 0 else "操作失败，请把日志反馈给分发人。") + "\n")
                self.outq.put(("__DONE__", code))
            except Exception as e:  # noqa: BLE001
                self.outq.put(f"\n[异常] {e}\n")
                self.outq.put(("__DONE__", 1))

        self.worker = threading.Thread(target=wrap, daemon=True)
        self.worker.start()

    def act_install(self) -> None:
        self._start(lambda: Runner(self.outq).run(ps("scripts\\setup.ps1", ["-NonInteractive"]), payload_dir()))

    def act_uninstall(self) -> None:
        self._start(lambda: Runner(self.outq).run(ps("scripts\\setup.ps1", ["-Uninstall", "-NonInteractive"]), payload_dir()))

    def act_elevated(self) -> None:
        self._start(lambda: Runner(self.outq).run_elevated(
            os.path.join(payload_dir(), "run-elevated.ps1"), "build\\install-elevated.ps1"))

    def act_uninstall_elevated(self) -> None:
        self._start(lambda: Runner(self.outq).run_elevated(
            os.path.join(payload_dir(), "run-elevated.ps1"), "build\\install-elevated.ps1 -Uninstall"))

    def _drain_queue(self) -> None:
        try:
            while True:
                item = self.outq.get_nowait()
                if isinstance(item, tuple) and item and item[0] == "__DONE__":
                    for b in self.buttons.values():
                        b.configure(state="normal")
                    self.status.configure(text="完成" if item[1] == 0 else "失败（见日志）")
                else:
                    self.log(item)
        except queue.Empty:
            pass
        self.root.after(100, self._drain_queue)

    def run(self) -> None:
        self.root.mainloop()


# ---------------- 静默模式 ----------------
def run_silent(mode: str) -> int:
    pdir = payload_dir()
    logf = log_file_path()
    append_log_file(f"===== AttachLinkSetup {mode} @ {__import__('time').strftime('%Y-%m-%d %H:%M:%S')} =====")

    def emit(line: str) -> None:
        append_log_file(line.rstrip("\n"))

    if mode == "/selfcheck":
        need = ["VERSION.txt", "dist\\AttachLink.xlam", "scripts\\setup.ps1", "scripts\\watchdog.ps1",
                "scripts\\watchdog_launcher.vbs", "build\\install.ps1", "build\\install-elevated.ps1",
                "run-elevated.ps1", "elevated-capture.ps1"]
        missing = [n for n in need if not os.path.exists(os.path.join(pdir, n))]
        print(f"AttachLink v{read_version()}  payload={pdir}")
        for n in need:
            print(("  OK  " if n not in missing else "  缺失 ") + n)
        return 0 if not missing else 3

    if mode == "/install":
        cmd = ps("scripts\\setup.ps1", ["-NonInteractive"])
    elif mode == "/uninstall":
        cmd = ps("scripts\\setup.ps1", ["-Uninstall", "-NonInteractive"])
    elif mode in ("/deploy-elevated", "/uninstall-elevated"):
        target = "build\\install-elevated.ps1" + (" -Uninstall" if mode == "/uninstall-elevated" else "")
        cmd = ps("run-elevated.ps1", [target])
    else:
        print(f"未知参数：{mode}")
        return 1

    proc = subprocess.run(cmd, cwd=pdir, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")
    emit(proc.stdout)
    if proc.stderr:
        emit("[stderr] " + proc.stderr)
    print(f"[AttachLinkSetup] 退出码 {proc.returncode}；完整日志：{logf}")
    return proc.returncode


def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1].startswith("/"):
        sys.exit(run_silent(sys.argv[1].lower()))
    App().run()


if __name__ == "__main__":
    main()

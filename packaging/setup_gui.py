# -*- coding: utf-8 -*-
"""
AttachLink 安装器（GUI 外壳）
- 业务逻辑全部走内嵌的 PowerShell 脚本（setup.ps1 / install-elevated.ps1），
  本程序只负责界面、日志展示与提权桥接，不复制业务逻辑（单一事实源）。
- 静默参数（供自动化验证/高级用户）：
    AttachLinkSetup.exe /install             标准安装（用户级）
    AttachLinkSetup.exe /deploy-elevated     鼨信诺管理员环境部署（弹 UAC）
    AttachLinkSetup.exe /uninstall           一键卸载（用户级 + 机器级看门狗；检测到机器级时弹一次 UAC）
    AttachLinkSetup.exe /uninstall-elevated  仅卸载机器级（弹 UAC，兼容保留）
    AttachLinkSetup.exe /selfcheck           自检（校验内嵌文件完整性）
  退出码：0 成功；2 Excel 在运行；3 安装包不完整；8 卸载未完全成功（机器级残留或 UAC 取消）；9 加载项列表清理未完全成功；1 其他失败。
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


def exe_dir() -> str:
    """安装器 exe 所在目录（分发 zip 的解压文件夹）"""
    if hasattr(sys, "frozen"):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def save_log_beside_exe(text: str, extra_note: str = "") -> str:
    """把日志快照另存到安装器所在目录（AttachLinkSetup.log，覆盖写），返回目标路径。
    目标不可写时静默失败返回空串（提示回落到 %TEMP% 路径）。"""
    target = os.path.join(exe_dir(), "AttachLinkSetup.log")
    try:
        with open(target, "w", encoding="utf-8") as f:
            f.write(text)
            if extra_note:
                f.write("\n" + extra_note + "\n")
        return target
    except OSError:
        return ""


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
        self.log(f"AttachLink v{self.ver}　|　每次操作结束后，日志快照会另存到本安装器所在文件夹（AttachLinkSetup.log），便于传递排查")

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
            ("uninst", "一键卸载（用户级 + 鼎信诺看门狗）", self.act_uninstall),
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
        # 一键卸载：setup.ps1 -Uninstall 自动检测机器级看门狗并弹 UAC 一并移除
        self._start(lambda: Runner(self.outq).run(ps("scripts\\setup.ps1", ["-Uninstall", "-NonInteractive"]), payload_dir()))

    def act_elevated(self) -> None:
        self._start(lambda: Runner(self.outq).run_elevated(
            os.path.join(payload_dir(), "run-elevated.ps1"), "build\\install-elevated.ps1"))

    def _drain_queue(self) -> None:
        try:
            while True:
                item = self.outq.get_nowait()
                if isinstance(item, tuple) and item and item[0] == "__DONE__":
                    code = item[1]
                    for b in self.buttons.values():
                        b.configure(state="normal")
                    # 日志快照另存到安装器所在目录，便于同事直接找到并转发
                    beside = save_log_beside_exe(self.logbox.get("1.0", "end-1c"))
                    if code == 0:
                        self.status.configure(text="✓ 操作成功——安装/卸载已完成", foreground="#1a7f37")
                        self.log("\n══════════ 操作成功 ══════════\n")
                        import tkinter.messagebox as mb
                        mb.showinfo(
                            "AttachLink 安装器",
                            "操作成功，安装（卸载）已完成。\n\n"
                            "请【完全退出并重新打开 Excel】后使用。\n"
                            "（编制机：管理员鼎信诺内使用还需另行执行\n"
                            "「鼎信诺管理员环境部署」）",
                            parent=self.root,
                        )
                    else:
                        hint = {2: "请先关闭所有 Excel（含鼎信诺）后重试", 3: "安装包不完整，请联系分发人",
                                8: "卸载未完全成功（机器级看门狗残留或 UAC 被取消），详见日志"}.get(code, "")
                        where = ("日志已保存到安装器所在文件夹：\n" + beside if beside
                                 else "日志文件（%TEMP%）：\n" + log_file_path())
                        self.status.configure(text="✗ 操作失败（退出码 %d）——日志已另存到安装目录" % code, foreground="#c62828")
                        self.log("\n══════════ 操作失败（退出码 %d）══════════\n日志快照：%s\n" % (code, beside or log_file_path()))
                        import tkinter.messagebox as mb
                        mb.showerror(
                            "AttachLink 安装器",
                            "操作失败（退出码 %d）。\n%s\n\n%s\n请把该文件发送给分发人排查。"
                            % (code, hint + ("\n" if hint else ""), where),
                            parent=self.root,
                        )
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
    # 日志快照另存到 exe 所在目录（方便传递给分发人；目标不可写则跳过）
    try:
        with open(logf, encoding="utf-8") as f:
            full_log = f.read()
    except OSError:
        full_log = proc.stdout
    beside = save_log_beside_exe(full_log, f"[silent {mode} exit={proc.returncode}]")
    if beside:
        print(f"[AttachLinkSetup] 日志快照已另存：{beside}")
    return proc.returncode


def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1].startswith("/"):
        sys.exit(run_silent(sys.argv[1].lower()))
    App().run()


if __name__ == "__main__":
    main()

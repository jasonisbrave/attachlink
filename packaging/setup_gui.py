# -*- coding: utf-8 -*-
"""
AttachLink 安装器的 tkinter 界面：只负责控件、日志展示与后台动作调度。
安装、卸载、结果文案和静默参数均见 setup_actions.py，本文件不复制业务逻辑。
"""
import queue
import sys
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

from setup_actions import (
    APP_TITLE,
    OK,
    PARTIAL,
    ActionResult,
    append_log_file,
    install_all,
    log_file_path,
    read_version,
    run_cli,
    safe,
    save_log_beside_exe,
    standard_install,
    uninstall,
)


class App:
    def __init__(self) -> None:
        self.ver = read_version()
        self.outq: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.more_visible = False
        self.buttons: dict[str, tk.Button | ttk.Button] = {}

        self.root = tk.Tk()
        self.root.title(f"{APP_TITLE} v{self.ver}")
        self.root.geometry("820x600")
        self.root.minsize(700, 520)
        self._build_ui()
        self.root.after(100, self._drain_queue)
        self.log(f"AttachLink v{self.ver}　|　每次操作结束后，日志快照会另存到本安装器所在文件夹（AttachLinkSetup.log），便于传递排查")

    # ---------------- UI ----------------
    def _build_ui(self) -> None:
        style = ttk.Style(self.root)
        style.configure(".", font=("Microsoft YaHei UI", 9))
        style.configure("Uninstall.TButton", padding=(20, 12))
        style.configure("Disclosure.TButton", padding=(4, 6))
        style.configure("Secondary.TButton", padding=(16, 8))

        head = ttk.Frame(self.root, padding=(20, 18, 20, 0))
        head.pack(fill="x")
        ttk.Label(
            head,
            text=f"AttachLink 审计底稿附件链接加载项　v{self.ver}",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            head,
            text="一键安装 = 启用加载项 + 部署鼎信诺环境（中途会弹一次管理员授权，点「是」）。",
            foreground="#4b5563",
            padding=(0, 10, 0, 0),
        ).pack(anchor="w")
        ttk.Label(
            head,
            text="不用鼎信诺，或没有管理员权限？展开下方「更多选项」，选「仅安装加载项」。",
            foreground="#4b5563",
            padding=(0, 5, 0, 0),
        ).pack(anchor="w")

        main_actions = ttk.Frame(self.root, padding=(20, 18, 20, 12))
        main_actions.pack(fill="x")
        self.buttons["install_all"] = tk.Button(
            main_actions,
            text="一键安装",
            font=("Microsoft YaHei UI", 11, "bold"),
            foreground="#ffffff",
            background="#176b4f",
            activeforeground="#ffffff",
            activebackground="#125a42",
            disabledforeground="#e1e8e5",
            relief="flat",
            borderwidth=0,
            padx=24,
            pady=10,
            takefocus=True,
            default="active",
            command=self.act_install_all,
        )
        self.buttons["install_all"].pack(side="left", padx=(0, 10))
        self.buttons["uninstall"] = ttk.Button(
            main_actions,
            text="一键卸载",
            style="Uninstall.TButton",
            command=self.act_uninstall,
        )
        self.buttons["uninstall"].pack(side="left")

        self.logbox = scrolledtext.ScrolledText(
            self.root,
            height=12,
            wrap="word",
            state="disabled",
            font=("Consolas", 10),
            background="#fafafa",
            foreground="#263238",
            relief="flat",
            highlightthickness=1,
            highlightbackground="#d5dae0",
            padx=10,
            pady=8,
        )
        self.logbox.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        more_area = ttk.Frame(self.root, padding=(20, 0, 20, 8))
        more_area.pack(fill="x")
        self.buttons["more"] = ttk.Button(
            more_area,
            text="▸  更多选项",
            style="Disclosure.TButton",
            command=self.toggle_more,
        )
        self.buttons["more"].pack(anchor="w")

        self.more_panel = ttk.Frame(more_area, padding=(20, 8, 0, 2))
        self.buttons["standard"] = ttk.Button(
            self.more_panel,
            text="仅安装加载项",
            style="Secondary.TButton",
            command=self.act_standard_install,
        )
        self.buttons["standard"].pack(side="left")

        self.status = ttk.Label(
            self.root,
            text="就绪",
            foreground="#4b5563",
            padding=(20, 8, 20, 10),
            relief="sunken",
            anchor="w",
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        self.status.pack(fill="x")

    def toggle_more(self) -> None:
        """通过显隐同一面板折叠单项操作；按钮本身可用 Tab 聚焦、空格触发。"""
        self.more_visible = not self.more_visible
        if self.more_visible:
            self.more_panel.pack(fill="x")
            self.buttons["more"].configure(text="▾  更多选项")
        else:
            self.more_panel.pack_forget()
            self.buttons["more"].configure(text="▸  更多选项")

    def log(self, text: str) -> None:
        append_log_file(text.rstrip("\n"))
        self.logbox.configure(state="normal")
        self.logbox.insert("end", text if text.endswith("\n") else text + "\n")
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    # ---------------- actions ----------------
    def _set_buttons_state(self, state: str) -> None:
        for button in self.buttons.values():
            button.configure(state=state)
        primary_color = "#176b4f" if state == "normal" else "#a5b5ae"
        self.buttons["install_all"].configure(background=primary_color)

    def _start(self, action) -> None:
        if self.worker and self.worker.is_alive():
            return
        self._set_buttons_state("disabled")
        self.status.configure(text="执行中…", foreground="#4b5563")

        def wrap() -> None:
            result = safe(action, self.outq)
            self.outq.put(("__DONE__", result))

        self.worker = threading.Thread(target=wrap, daemon=True)
        self.worker.start()

    def act_install_all(self) -> None:
        self._start(install_all)

    def act_uninstall(self) -> None:
        self._start(uninstall)

    def act_standard_install(self) -> None:
        self._start(standard_install)

    def _finish(self, result: ActionResult) -> None:
        self._set_buttons_state("normal")

        beside = save_log_beside_exe(self.logbox.get("1.0", "end-1c"))
        log_where = (
            "日志已保存到安装器所在文件夹：\n" + beside
            if beside
            else "日志文件（%TEMP%）：\n" + log_file_path()
        )
        dialog_text = result.dialog_text(log_where)

        if result.kind == OK:
            self.status.configure(text="✓ " + result.title, foreground="#1a7f37")
            messagebox.showinfo(APP_TITLE, dialog_text, parent=self.root)
        elif result.kind == PARTIAL:
            self.status.configure(text="⚠ " + result.title, foreground="#9a6700")
            messagebox.showwarning(APP_TITLE, dialog_text, parent=self.root)
        else:
            self.status.configure(text="✗ " + result.title, foreground="#c62828")
            messagebox.showerror(APP_TITLE, dialog_text, parent=self.root)

    def _drain_queue(self) -> None:
        try:
            while True:
                item = self.outq.get_nowait()
                if isinstance(item, tuple) and item and item[0] == "__DONE__":
                    self._finish(item[1])
                else:
                    self.log(item)
        except queue.Empty:
            pass
        self.root.after(100, self._drain_queue)

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1].startswith("/"):
        sys.exit(run_cli(sys.argv[1].lower()))
    App().run()


if __name__ == "__main__":
    main()

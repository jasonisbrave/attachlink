# -*- coding: utf-8 -*-
"""
AttachLink 安装器的 tkinter 界面：只负责控件、日志展示与后台动作调度。
安装、卸载、检查更新、一键更新的动作、结果文案和静默参数均见 setup_actions.py，本文件不复制业务逻辑。
"""
import queue
import re
import sys
import threading
import tkinter as tk
import webbrowser
from tkinter import messagebox, scrolledtext, ttk

from setup_actions import (
    APP_TITLE,
    OK,
    PARTIAL,
    ActionResult,
    append_log_file,
    apply_update,
    check_update,
    exe_dir,
    handle_update_handoff,
    handoff_summary_line,
    health_check,
    install_all,
    log_file_path,
    read_version,
    run_cli,
    safe,
    save_log_beside_exe,
    standard_install,
    uninstall,
    _log_beside_dir,
)

# 界面用色令牌：绿色主按钮、状态三色（✓ ⚠ ✗）、灰阶说明文字、更新提示条的琥珀底色
GREEN = "#176b4f"
GREEN_ACTIVE = "#125a42"
GREEN_DISABLED = "#a5b5ae"
TEXT_MUTED = "#4b5563"
STATUS_OK = "#1a7f37"
STATUS_WARN = "#9a6700"
STATUS_FAIL = "#c62828"
BANNER_BG = "#faf3e1"
BANNER_BORDER = "#e3d0a0"
BANNER_TEXT = "#7a5d00"
LINK_BLUE = "#0b5cad"
PROGRESS_PREFIX = "【进度】"   # apply_update 推送的进度行前缀：`【进度】已下载\t总大小\t百分比%`


def save_gui_log_snapshot(text: str, extra_note: str = "") -> str:
    """界面收尾统一的日志快照另存：交接模式（--updated-from）写到旧安装器所在目录——
    新安装器由交接从 %TEMP% 更新临时目录启动，快照写在那里用户看不到；无交接参数或旧目录
    不存在时回落 exe 所在目录。不依赖 Tk 窗口，三个收尾路径都走这里，避免调用点漂移。"""
    return save_log_beside_exe(text, extra_note, target_dir=_log_beside_dir())


class App:
    def __init__(self, handoff: dict | None = None, autostart_install: bool = False) -> None:
        self.ver = read_version()
        self.outq: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.more_visible = False
        self.buttons: dict[str, tk.Button | ttk.Button] = {}
        # 当前动作名（check_update / apply_update / 安装卸载动作），收尾时按动作分流呈现方式
        self.current_action = ""
        # 检查到的新版本信息与发布说明地址（结果对象尚未带回 notes_url，拿到之前链接不显示）
        self.update_info: dict | None = None
        self.update_notes_url: str | None = None
        self.progress_anim = False   # 总大小未知时进度条的流动动画是否已开启

        self.root = tk.Tk()
        self.root.title(f"{APP_TITLE} v{self.ver}")
        self.root.geometry("820x600")
        self.root.minsize(700, 520)
        # 关窗守卫：后台任务（安装 / 卸载 / 更新）还活着时忽略关闭——半途掐断会留下半成品状态；
        # 已知限制：不做跨进程单实例互斥（会与一键更新交接冲突：新安装器是在旧安装器仍存活时启动的）
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()
        self.root.after(100, self._drain_queue)
        # 交接结果如实呈现：替换失败（等待超时 / 杀毒锁定 / 目录不可写）时旧 exe 仍是旧版，
        # 必须告知用户本安装器所在位置，而不是无条件说「已更新」
        handoff_line = handoff_summary_line(handoff, exe_dir())
        if handoff_line:
            self.log(handoff_line)
        self.log(f"AttachLink v{self.ver}　|　每次操作结束后，日志快照会另存到本安装器所在文件夹（AttachLinkSetup.log），便于传递排查")
        if autostart_install:
            # 交接安装：窗口出现约 300ms 后自动执行一次与「一键安装」按钮完全相同的动作
            self.root.after(300, self.act_install_all)

    # ---------------- UI ----------------
    def _build_ui(self) -> None:
        style = ttk.Style(self.root)
        style.configure(".", font=("Microsoft YaHei UI", 9))
        style.configure("Uninstall.TButton", padding=(20, 12))
        style.configure("Disclosure.TButton", padding=(4, 6))
        style.configure("Secondary.TButton", padding=(16, 8))
        style.configure("Check.TButton", padding=(14, 4))

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
            foreground=TEXT_MUTED,
            padding=(0, 10, 0, 0),
        ).pack(anchor="w")
        ttk.Label(
            head,
            text="不用鼎信诺，或没有管理员权限？展开下方「更多选项」，选「仅安装加载项」。",
            foreground=TEXT_MUTED,
            padding=(0, 5, 0, 0),
        ).pack(anchor="w")

        main_actions = ttk.Frame(self.root, padding=(20, 18, 20, 12))
        main_actions.pack(fill="x")
        self.buttons["install_all"] = tk.Button(
            main_actions,
            text="一键安装",
            font=("Microsoft YaHei UI", 11, "bold"),
            foreground="#ffffff",
            background=GREEN,
            activeforeground="#ffffff",
            activebackground=GREEN_ACTIVE,
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
        # 健康检查结果行按前缀着色（前缀约定见 setup_actions.health_check；↳ 提示行与所在项同前缀同色）
        self.logbox.tag_configure("hc-ok", foreground="#1a7f37")
        self.logbox.tag_configure("hc-warn", foreground="#9a6700")
        self.logbox.tag_configure("hc-fail", foreground="#c62828", font=("Consolas", 10, "bold"))
        self.logbox.tag_configure("hc-skip", foreground="#8a9099")
        self.logbox.tag_configure("hc-info", foreground="#4a6785")

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
        self.buttons["health"] = ttk.Button(
            self.more_panel,
            text="健康检查",
            style="Secondary.TButton",
            command=self.act_health_check,
        )
        self.buttons["health"].pack(side="left", padx=(10, 0))
        ttk.Label(
            self.more_panel,
            text="只读检查两级部署、启用登记与看门狗心跳，不改任何设置",
            foreground="#8a9099",
            font=("Microsoft YaHei UI", 9),
        ).pack(side="left", padx=(12, 0))

        # —— 更新提示条 / 下载进度条：默认不显示，检查到新版本时插到日志框上方（高度约 40px）——
        # 本机 vista 主题下 ttk 控件自定义配色不生效：需要配色的按钮一律用 tk.Button 显式配色
        banner = tk.Frame(
            self.root,
            background=BANNER_BG,
            highlightthickness=1,
            highlightbackground=BANNER_BORDER,
        )
        self.update_banner = banner

        self.banner_msg = tk.Frame(banner, background=BANNER_BG)
        self.banner_msg.pack(fill="x")
        self.banner_label = tk.Label(
            self.banner_msg,
            text="",
            anchor="w",
            background=BANNER_BG,
            foreground=BANNER_TEXT,
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        self.banner_label.pack(side="left", padx=(12, 6), pady=9)
        self.banner_notes = tk.Label(
            self.banner_msg,
            text="发布说明",
            background=BANNER_BG,
            foreground=LINK_BLUE,
            font=("Microsoft YaHei UI", 9, "underline"),
            cursor="hand2",
        )
        self.banner_notes.bind("<Button-1>", self._open_notes)
        self.buttons["update_now"] = tk.Button(
            self.banner_msg,
            text="一键更新",
            font=("Microsoft YaHei UI", 9, "bold"),
            foreground="#ffffff",
            background=GREEN,
            activeforeground="#ffffff",
            activebackground=GREEN_ACTIVE,
            disabledforeground="#dce7e2",
            relief="flat",
            borderwidth=0,
            padx=16,
            pady=4,
            takefocus=True,
            command=self.act_apply_update,
        )
        self.buttons["update_now"].pack(side="right", padx=(8, 12), pady=6)

        self.banner_progress = tk.Frame(banner, background=BANNER_BG)
        self.progressbar = ttk.Progressbar(
            self.banner_progress,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            value=0,
        )
        self.progressbar.pack(side="left", fill="x", expand=True, padx=(12, 10), pady=11)
        self.progress_label = tk.Label(
            self.banner_progress,
            text="",
            anchor="e",
            background=BANNER_BG,
            foreground=TEXT_MUTED,
            font=("Microsoft YaHei UI", 9),
        )
        self.progress_label.pack(side="right", padx=(0, 12))

        # —— 底部：状态栏 + 右侧安静的「检查更新」次级入口 ——
        statusbar = ttk.Frame(self.root)
        statusbar.pack(fill="x", side="bottom")
        self.buttons["check_update"] = ttk.Button(
            statusbar,
            text="检查更新",
            style="Check.TButton",
            command=self.act_check_update,
        )
        self.buttons["check_update"].pack(side="right", padx=(8, 12), pady=6)
        self.status = ttk.Label(
            statusbar,
            text="就绪",
            foreground=TEXT_MUTED,
            padding=(20, 8, 8, 10),
            relief="sunken",
            anchor="w",
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        self.status.pack(side="left", fill="both", expand=True)

    def toggle_more(self) -> None:
        """通过显隐同一面板折叠单项操作；按钮本身可用 Tab 聚焦、空格触发。"""
        self.more_visible = not self.more_visible
        if self.more_visible:
            self.more_panel.pack(fill="x")
            self.buttons["more"].configure(text="▾  更多选项")
        else:
            self.more_panel.pack_forget()
            self.buttons["more"].configure(text="▸  更多选项")

    @staticmethod
    def _hc_tag(text: str) -> str:
        """健康检查结果行（以【通过】等五种前缀开头，含 ↳ 提示行）返回着色 tag，其余日志不着色。"""
        for prefix, tag in (
            ("【异常】", "hc-fail"),
            ("【留意】", "hc-warn"),
            ("【通过】", "hc-ok"),
            ("【跳过】", "hc-skip"),
            ("【信息】", "hc-info"),
        ):
            if text.startswith(prefix):
                return tag
        return ""

    def log(self, text: str) -> None:
        append_log_file(text.rstrip("\n"))
        chunk = text if text.endswith("\n") else text + "\n"
        tag = self._hc_tag(text)
        self.logbox.configure(state="normal")
        if tag:
            self.logbox.insert("end", chunk, tag)
        else:
            self.logbox.insert("end", chunk)
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    # ---------------- 更新提示条 ----------------
    def _show_banner(self, latest: str, current: str, notes_url: str | None = None) -> None:
        """检查到新版本：在日志框上方展示内联提示条（文字 + 一键更新 + 发布说明链接）"""
        self.update_info = {"latest": latest, "notes": notes_url}
        self.update_notes_url = notes_url
        self.banner_label.configure(text=f"发现新版本 v{latest}（当前 v{current}）")
        if notes_url:
            self.banner_notes.pack(side="left", padx=(0, 4), pady=9)
        else:
            self.banner_notes.pack_forget()
        self._reset_progress()
        self.banner_progress.pack_forget()
        self.banner_msg.pack(fill="x")
        self.update_banner.pack(fill="x", padx=20, pady=(0, 8), before=self.logbox)

    def _banner_to_progress(self) -> None:
        """开始更新：提示条转为进度状态（收起文字与按钮，换上进度条）"""
        self.banner_msg.pack_forget()
        self._reset_progress()
        self.progress_label.configure(text="准备下载…")
        self.banner_progress.pack(fill="x")
        if not self.update_banner.winfo_manager():
            self.update_banner.pack(fill="x", padx=20, pady=(0, 8), before=self.logbox)

    def _hide_banner(self) -> None:
        self.update_banner.pack_forget()
        self.banner_progress.pack_forget()
        self.banner_msg.pack(fill="x")
        self.update_info = None
        self.update_notes_url = None
        self._reset_progress()

    def _reset_progress(self) -> None:
        if self.progress_anim:
            self.progressbar.stop()
            self.progress_anim = False
        self.progressbar.configure(mode="determinate", maximum=100, value=0)

    def _apply_progress(self, line: str) -> None:
        """解析后台推送的进度行，只刷新进度条与文字（进度行不进日志框）"""
        body = line.strip()
        if body.endswith("%"):
            body = body[:-1]
        fields = body[len(PROGRESS_PREFIX):].split("\t")
        if len(fields) < 2:
            return
        try:
            done = max(0, int(fields[0].strip()))
            total = int(fields[1].strip())
        except ValueError:
            return
        if total and total > 0:
            if self.progress_anim:
                self.progressbar.stop()
                self.progress_anim = False
            self.progressbar.configure(mode="determinate", maximum=total, value=min(done, total))
            pct = min(100, done * 100 // total)
            self.progress_label.configure(
                text=f"已下载 {self._fmt_size(done)} / {self._fmt_size(total)}（{pct}%）")
        else:
            # 总大小未知：进度条转流动动画，文字只报已下载量
            if not self.progress_anim:
                self.progressbar.configure(mode="indeterminate", maximum=100, value=0)
                self.progressbar.start(60)
                self.progress_anim = True
            self.progress_label.configure(text=f"已下载 {self._fmt_size(done)}")

    @staticmethod
    def _fmt_size(n: float) -> str:
        if n >= 1024 * 1024:
            return f"{n / (1024 * 1024):.1f} MB"
        if n >= 1024:
            return f"{n / 1024:.0f} KB"
        return f"{n:.0f} B"

    @staticmethod
    def _parse_latest(title: str) -> str:
        """从结果标题（如「发现新版本 v9.9.9」）里取版本号，取不到时占位为 ?"""
        matched = re.search(r"v(\d+(?:\.\d+)+)", title)
        return matched.group(1) if matched else "?"

    def _open_notes(self, _event=None) -> None:
        if self.update_notes_url:
            webbrowser.open(self.update_notes_url)

    def _on_close(self) -> None:
        # 关窗守卫：后台动作还在跑时忽略关闭请求（动作完成或失败后会自行恢复按钮并可正常关闭）
        if self.worker and self.worker.is_alive():
            messagebox.showwarning(
                APP_TITLE,
                "安装 / 卸载 / 更新正在进行，请等待完成后再关闭。",
                parent=self.root,
            )
            return
        self.root.destroy()

    # ---------------- actions ----------------
    def _set_buttons_state(self, state: str) -> None:
        for button in self.buttons.values():
            button.configure(state=state)
        primary_color = GREEN if state == "normal" else GREEN_DISABLED
        for key in ("install_all", "update_now"):
            self.buttons[key].configure(background=primary_color)

    def _start(self, action, name: str = "", status_text: str = "执行中…") -> None:
        if self.worker and self.worker.is_alive():
            return
        self.current_action = name
        self._set_buttons_state("disabled")
        self.status.configure(text=status_text, foreground=TEXT_MUTED)

        def wrap() -> None:
            result = safe(action, self.outq)
            self.outq.put(("__DONE__", result))

        self.worker = threading.Thread(target=wrap, daemon=True)
        self.worker.start()

    def act_install_all(self) -> None:
        self._start(install_all, name="install_all")

    def act_uninstall(self) -> None:
        self._start(uninstall, name="uninstall")

    def act_standard_install(self) -> None:
        self._start(standard_install, name="standard_install")

    def act_check_update(self) -> None:
        self._start(check_update, name="check_update", status_text="正在检查更新…")

    def act_apply_update(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        latest = (self.update_info or {}).get("latest") or "?"
        confirmed = messagebox.askyesno(
            APP_TITLE,
            f"将下载并安装新版本 v{latest}。\n\n"
            "开始前请先关闭 Excel（含鼎信诺）。\n\n"
            "过程中会弹一次管理员授权，请点「是」。",
            parent=self.root,
        )
        if not confirmed:
            return
        self._banner_to_progress()
        # 交接动作固定带 --autostart-install：新版本安装器打开自己的窗口并自动开始安装
        self._start(lambda q: apply_update(q, action="--autostart-install"),
                    name="apply_update", status_text="正在更新…")

    def act_health_check(self) -> None:
        # name/status_text 必须关键字传参：_start 的第二个位置参数是动作名（供 _finish 分流），不是状态栏文案
        self._start(health_check, name="health_check", status_text="正在健康检查…（只读，约几秒）")

    def _finish(self, result: ActionResult) -> None:
        action, self.current_action = self.current_action, ""
        self._set_buttons_state("normal")
        if action == "check_update":
            self._finish_check(result)
        elif action == "apply_update":
            self._finish_update(result)
        else:
            self._finish_dialog(result)

    def _finish_check(self, result: ActionResult) -> None:
        """检查更新收尾：三种结果都不弹模态框——没事的事不打扰用户，原因都看日志"""
        if result.kind == OK:
            self._hide_banner()
            self.status.configure(text=f"✓ {result.title} v{self.ver}", foreground=STATUS_OK)
        elif result.kind == PARTIAL:
            self.status.configure(text="⚠ " + result.title, foreground=STATUS_WARN)
            # 发布说明地址：后端在 body 里以「发布说明：」单独一行给出（清单没有则没有这一行）
            notes_url = None
            for line in (result.body or "").splitlines():
                if line.startswith("发布说明："):
                    notes_url = line[len("发布说明："):].strip() or None
            self._show_banner(latest=self._parse_latest(result.title), current=self.ver, notes_url=notes_url)
        else:
            self._hide_banner()
            self.status.configure(text="⚠ " + result.title, foreground=STATUS_WARN)

    def _finish_update(self, result: ActionResult) -> None:
        self._hide_banner()
        if result.hint == "EXIT_AFTER":
            # 交接成功：日志快照照常保存，不弹任何框直接关窗；新版本安装器的窗口接着出现
            save_gui_log_snapshot(self.logbox.get("1.0", "end-1c"))
            self.status.configure(text="✓ " + result.title, foreground=STATUS_OK)
            self.root.destroy()
            return
        if result.kind == OK:
            # 更新流程里发现已是最新：与检查更新同样处理，只更新状态栏不弹框
            self.status.configure(text=f"✓ {result.title} v{self.ver}", foreground=STATUS_OK)
            return
        beside = save_gui_log_snapshot(self.logbox.get("1.0", "end-1c"))
        log_where = (
            "日志已保存到安装器所在文件夹：\n" + beside
            if beside
            else "日志文件（%TEMP%）：\n" + log_file_path()
        )
        self.status.configure(text="✗ " + result.title, foreground=STATUS_FAIL)
        messagebox.showerror(APP_TITLE, result.dialog_text(log_where), parent=self.root)

    def _finish_dialog(self, result: ActionResult) -> None:
        beside = save_gui_log_snapshot(self.logbox.get("1.0", "end-1c"))
        log_where = (
            "日志已保存到安装器所在文件夹：\n" + beside
            if beside
            else "日志文件（%TEMP%）：\n" + log_file_path()
        )
        dialog_text = result.dialog_text(log_where)

        if result.kind == OK:
            self.status.configure(text="✓ " + result.title, foreground=STATUS_OK)
            messagebox.showinfo(APP_TITLE, dialog_text, parent=self.root)
        elif result.kind == PARTIAL:
            self.status.configure(text="⚠ " + result.title, foreground=STATUS_WARN)
            messagebox.showwarning(APP_TITLE, dialog_text, parent=self.root)
        else:
            self.status.configure(text="✗ " + result.title, foreground=STATUS_FAIL)
            messagebox.showerror(APP_TITLE, dialog_text, parent=self.root)

    def _drain_queue(self) -> None:
        finished: ActionResult | None = None
        progress_line: str | None = None
        try:
            # 一轮把队列里现存的行全部取完；进度行很密，只留最后一帧渲染，避免积压
            while True:
                item = self.outq.get_nowait()
                if isinstance(item, tuple) and item and item[0] == "__DONE__":
                    finished = item[1]
                elif isinstance(item, str) and item.startswith(PROGRESS_PREFIX):
                    progress_line = item
                else:
                    self.log(item)
        except queue.Empty:
            pass
        if progress_line:
            self._apply_progress(progress_line)
        if finished is not None:
            self._finish(finished)
        self.root.after(100, self._drain_queue)

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1].startswith("/"):
        sys.exit(run_cli(sys.argv[1].lower()))
    hand = handle_update_handoff()   # 更新交接：命令行带 --updated-from 时先替换旧安装器文件，再进界面
    args = sys.argv[1:]
    autostart = "--autostart-install" in args
    App(handoff=hand, autostart_install=autostart).run()


if __name__ == "__main__":
    main()

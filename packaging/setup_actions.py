# -*- coding: utf-8 -*-
"""
AttachLink 安装器 —— 动作层（v1.4.4）
- 业务逻辑全部走内嵌的 PowerShell 脚本（setup.ps1 / install-elevated.ps1），本模块只负责
  编排、退出码解读与结果文案，不含任何界面代码；界面见 setup_gui.py（不复制业务逻辑，单一事实源）。
- 动作（均接收一个带 .put(line) 的队列对象，逐行推送日志；返回 ActionResult）：
    install_all        一键安装 = standard_install 成功后再 deploy_elevated
    standard_install   仅安装加载项（用户级）
    deploy_elevated    仅部署鼎信诺环境（机器级看门狗，弹 UAC）——界面不提供单独入口，只供 install_all
                       内部调用与 /deploy-elevated 自动化维护用；只装过加载项的机器重新「一键安装」覆盖即可
    uninstall          一键卸载（用户级 + 机器级看门狗；检测到机器级时弹一次 UAC）
- 静默参数（供自动化验证 / 高级用户）：
    AttachLinkSetup.exe /install-all        一键安装（= /install 成功后接着 /deploy-elevated）
    AttachLinkSetup.exe /install            仅安装加载项（标准安装）
    AttachLinkSetup.exe /deploy-elevated    仅部署鼎信诺环境（弹 UAC；自动化 / 维护用，界面无此入口）
    AttachLinkSetup.exe /uninstall          一键卸载（用户级 + 机器级看门狗；检测到机器级时弹一次 UAC）
    AttachLinkSetup.exe /uninstall-elevated 仅卸载机器级（弹 UAC，兼容保留）
    AttachLinkSetup.exe /selfcheck          自检（校验内嵌文件完整性）
- 退出码：0 成功；2 Excel 在运行；3 安装包不完整；8 卸载未完全成功（机器级残留或 UAC 取消）；
  9 加载项列表清理未完全成功；1223 UAC 被取消；1 其他失败。
  /install-all 任一步失败即返回该步退出码（第一步失败不会继续第二步；第二步失败时用户级已装好，
  结果为「部分完成」而非全失败）。
"""
import os
import subprocess
import sys
import time
from dataclasses import dataclass

VERSION_FILE = "VERSION.txt"   # 相对 payload_dir()；曾误写成 "payload\VERSION.txt" 导致读到 payload\payload\… 而显示「未知版本」
APP_TITLE = "AttachLink 安装器"

# 结果类别：成功 / 部分完成（用户级已装好，机器级未完成）/ 失败
OK, PARTIAL, FAIL = "ok", "partial", "fail"
UAC_CANCELLED = 1223

_HINTS = {
    2: "请先关闭所有 Excel（含鼎信诺）后重试",
    3: "安装包不完整，请联系分发人",
    8: "卸载未完全成功（机器级看门狗残留或 UAC 被取消），详见日志",
    9: "加载项列表登记未能完全清理，详见日志里的警告；若提示未能确认 Excel 已退出，关闭所有 Excel 后重新卸载即可",
}


# ---------------- 路径与日志文件 ----------------

def payload_dir() -> str:
    """PyInstaller onefile 运行时解包目录；源码直跑时用 packaging/payload"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, "payload")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "payload")


def read_version() -> str:
    try:
        with open(os.path.join(payload_dir(), VERSION_FILE), encoding="utf-8") as f:
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


def full_log_text() -> str:
    try:
        with open(log_file_path(), encoding="utf-8") as f:
            return f.read()
    except OSError:
        return ""


# ---------------- 结果对象 ----------------

@dataclass
class ActionResult:
    kind: str       # OK / PARTIAL / FAIL
    code: int       # 退出码（静默模式返回给调用方）
    title: str      # 状态栏一句话
    body: str       # 弹窗正文（不含日志位置）
    hint: str = ""  # 失败原因提示（仅失败类）

    def dialog_text(self, log_where: str) -> str:
        """弹窗全文：正文 +（失败原因）+ 日志位置"""
        if self.kind == OK:
            return self.body
        if self.kind == PARTIAL:
            return self.body + "\n\n" + log_where + "\n如需排查，请把该日志发送给分发人。"
        text = self.body
        if self.hint:
            text += "\n\n" + self.hint
        return text + "\n\n" + log_where + "\n请把该日志发送给分发人排查。"


def _failed(code: int, title: str, body: str) -> ActionResult:
    return ActionResult(FAIL, code, title, body, _HINTS.get(code, ""))


def safe(action, outq) -> ActionResult:
    """执行动作，未预期异常转成失败结果（调用方线程不抛异常）"""
    try:
        return action(outq)
    except Exception as e:  # noqa: BLE001
        outq.put(f"\n[异常] {e}\n")
        return ActionResult(FAIL, 1, "发生异常", f"执行过程中出现异常：{e}")


# ---------------- PowerShell 调用 ----------------

class Runner:
    """运行 PowerShell 子进程，stdout 逐行送入队列"""

    def __init__(self, outq) -> None:
        self.outq = outq

    def run(self, args: list, cwd: str) -> int:
        self.outq.put("$ " + " ".join(args[:4]) + " ...\n")
        proc = subprocess.Popen(
            args,
            cwd=cwd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        assert proc.stdout is not None
        for line in proc.stdout:
            self.outq.put(line)
        proc.stdout.close()
        return proc.wait()

    def run_elevated(self, target: str) -> int:
        """经 run-elevated.ps1 提权运行目标脚本：UAC 由系统弹出，日志回传临时文件"""
        return self.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
             os.path.join(payload_dir(), "run-elevated.ps1"), target],
            payload_dir(),
        )


def _ps(script_rel: str, extra: list | None = None) -> list:
    args = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
            os.path.join(payload_dir(), script_rel)]
    return args + (extra or [])


# ---------------- 动作 ----------------

def standard_install(outq, label: str = "") -> ActionResult:
    """用户级：复制 xlam → 受信任位置 → 注册启用加载项"""
    if label:
        outq.put("\n" + label + "\n")
    code = Runner(outq).run(_ps("scripts\\setup.ps1", ["-NonInteractive"]), payload_dir())
    if code == 0:
        return ActionResult(
            OK, 0, "加载项已启用",
            "加载项已安装并启用。\n\n"
            "请【完全退出并重新打开 Excel】后使用：在单元格上点右键，即可看到「插入附件链接（可多选）」。\n\n"
            "如需在本机以管理员运行的鼎信诺里编制底稿，还要部署鼎信诺环境：\n"
            "重新运行本安装器，点「一键安装」即可（会覆盖安装加载项并部署鼎信诺环境）。")
    return _failed(code, f"加载项安装失败（退出码 {code}）", "加载项未能安装。")


def deploy_elevated(outq, label: str = "") -> ActionResult:
    """机器级：Program Files 部署 xlam + 看门狗、HKLM 受信任位置、登录计划任务（弹一次 UAC）"""
    if label:
        outq.put("\n" + label + "\n")
    code = Runner(outq).run_elevated("build\\install-elevated.ps1")
    if code == 0:
        return ActionResult(
            OK, 0, "鼎信诺环境已部署",
            "机器级看门狗已部署并启动。\n\n"
            "请完全退出并重新打开鼎信诺和 Excel：在鼎信诺里打开底稿约 2~4 秒后，右键菜单自动出现。")
    if code == UAC_CANCELLED:
        return ActionResult(FAIL, code, "鼎信诺环境未部署（授权被取消）",
                            "管理员授权被取消，鼎信诺环境未部署。\n需要时请重新运行并允许授权。")
    return _failed(code, f"鼎信诺环境部署失败（退出码 {code}）", "机器级看门狗未能部署。")


def install_all(outq) -> ActionResult:
    """一键安装：先用户级（失败即停），再机器级（UAC 取消 / 失败 = 部分完成）"""
    first = standard_install(outq, "［1/2］安装加载项（用户级）")
    if first.kind != OK:
        return first
    second = deploy_elevated(outq, "［2/2］部署鼎信诺环境（将弹出管理员授权，请点「是」）")
    if second.kind == OK:
        return ActionResult(
            OK, 0, "一键安装完成",
            "加载项已启用，鼎信诺环境已部署。\n\n"
            "请【完全退出并重新打开 Excel（及鼎信诺）】后使用：\n"
            "· 在单元格上点右键，即可看到「插入附件链接（可多选）」\n"
            "· 在鼎信诺里打开底稿约 2~4 秒后，右键菜单自动出现")
    if second.code == UAC_CANCELLED:
        return ActionResult(
            PARTIAL, UAC_CANCELLED, "加载项已安装，鼎信诺环境未部署",
            "管理员授权被取消，鼎信诺环境未部署。\n\n"
            "· 不使用鼎信诺的同事：现在就可以使用，重新打开 Excel 即可。\n"
            "· 需要在鼎信诺里编制底稿的同事：请重新运行并允许授权（或请 IT 用管理员账户协助）。")
    return ActionResult(
        PARTIAL, second.code, "加载项已安装，鼎信诺环境部署失败",
        f"机器级看门狗未能部署（退出码 {second.code}）。\n\n"
        "· 日常 Excel 里的加载项功能不受影响，重新打开 Excel 即可使用。\n"
        "· 需要在鼎信诺里编制底稿的同事：请把日志反馈给分发人。")


def uninstall(outq) -> ActionResult:
    """一键卸载：用户级 + 机器级看门狗（setup.ps1 -Uninstall 自动检测并弹一次 UAC）"""
    outq.put("\n一键卸载（用户级加载项 + 鼎信诺看门狗）…\n")
    code = Runner(outq).run(_ps("scripts\\setup.ps1", ["-Uninstall", "-NonInteractive"]), payload_dir())
    if code == 0:
        return ActionResult(
            OK, 0, "卸载完成",
            "用户级加载项与机器级看门狗均已卸载。\n\n"
            "已打开的 Excel 会话里，右键菜单会保留到该会话关闭；重新打开 Excel 后不再出现。")
    if code == 9:
        return _failed(code, "卸载未完全成功（加载项列表清理未完全成功）",
                       "加载项已卸载，但加载项列表的登记未能完全清理。")
    if code == 8:
        return _failed(code, "卸载未完全成功", "卸载已执行，但机器级看门狗仍有残留，或管理员授权被取消。")
    return _failed(code, f"卸载失败（退出码 {code}）", "卸载未能完成。")


def uninstall_elevated_only(outq) -> ActionResult:
    """兼容保留：仅卸载机器级"""
    outq.put("\n仅卸载机器级看门狗…\n")
    code = Runner(outq).run_elevated("build\\install-elevated.ps1 -Uninstall")
    if code == 0:
        return ActionResult(OK, 0, "机器级看门狗已卸载", "机器级看门狗已卸载。用户级加载项未改动。")
    return _failed(code, f"机器级卸载失败（退出码 {code}）", "机器级看门狗未能卸载。")


# ---------------- 静默模式 ----------------

class _FileQueue:
    """静默模式用的队列替身：日志行直接落盘"""

    def put(self, line: str) -> None:
        append_log_file(line.rstrip("\n"))


def selfcheck() -> int:
    pdir = payload_dir()
    need = ["VERSION.txt", "dist\\AttachLink.xlam", "scripts\\setup.ps1", "scripts\\watchdog.ps1",
            "scripts\\watchdog_launcher.vbs", "build\\install.ps1", "build\\install-elevated.ps1",
            "run-elevated.ps1", "elevated-capture.ps1"]
    missing = [n for n in need if not os.path.exists(os.path.join(pdir, n))]
    ver = read_version()
    print(f"AttachLink v{ver}  payload={pdir}")
    for n in need:
        print(("  OK  " if n not in missing else "  缺失 ") + n)
    # 版本号读不到（文件在但内容为空 / 路径拼接错误）同样算自检失败——安装器标题、GUI 标题都依赖它
    if not ver or ver == "未知版本":
        print("  失败  版本号不可读（VERSION.txt 存在但读不到内容）")
        return 3
    return 0 if not missing else 3


CLI_ACTIONS = {
    "/install-all": install_all,
    "/install": standard_install,
    "/deploy-elevated": deploy_elevated,
    "/uninstall": uninstall,
    "/uninstall-elevated": uninstall_elevated_only,
}


def run_cli(mode: str) -> int:
    logf = log_file_path()
    append_log_file(f"===== AttachLinkSetup {mode} @ {time.strftime('%Y-%m-%d %H:%M:%S')} =====")
    if mode == "/selfcheck":
        return selfcheck()
    action = CLI_ACTIONS.get(mode)
    if action is None:
        print(f"未知参数：{mode}")
        return 1
    result = safe(action, _FileQueue())
    append_log_file(f"[结果] {result.title}（退出码 {result.code}）")
    print(f"[AttachLinkSetup] {result.title}；退出码 {result.code}；完整日志：{logf}")
    # 日志快照另存到 exe 所在目录（方便传递给分发人；目标不可写则跳过）
    beside = save_log_beside_exe(full_log_text(), f"[silent {mode} exit={result.code}] {result.title}")
    if beside:
        print(f"[AttachLinkSetup] 日志快照已另存：{beside}")
    return result.code

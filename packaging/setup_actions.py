# -*- coding: utf-8 -*-
r"""
AttachLink 安装器 —— 动作层（v1.6.0）
- 业务逻辑全部走内嵌的 PowerShell 脚本（setup.ps1 / install-elevated.ps1 / healthcheck.ps1），本模块只负责
  编排、退出码解读与结果文案，不含任何界面代码；界面见 setup_gui.py（不复制业务逻辑，单一事实源）。
- 动作（均接收一个带 .put(line) 的队列对象，逐行推送日志；返回 ActionResult）：
    install_all        一键安装 = standard_install 成功后再 deploy_elevated，最后自动 health_check
    standard_install   仅安装加载项（用户级）
    deploy_elevated    仅部署鼎信诺环境（机器级看门狗，弹 UAC）——界面不提供单独入口，只供 install_all
                       内部调用与 /deploy-elevated 自动化维护用；只装过加载项的机器重新「一键安装」覆盖即可
    uninstall          一键卸载（用户级 + 机器级看门狗；检测到机器级时弹一次 UAC）
    health_check       健康检查（纯只读：跑 scripts\healthcheck.ps1，逐条推送【通过】【留意】【异常】【跳过】
                       【信息】五种前缀的人读行；kind=ok/partial/fail，code=脚本退出码 0/10/1）
    check_update       检查更新（拉发布清单并与当前版本比较；不联网不安装）
    apply_update       一键更新（下载并校验新安装器 → 交接给新安装器安装；本进程随后退出）
- 静默参数（供自动化验证 / 高级用户）：
    AttachLinkSetup.exe /install-all        一键安装（= /install 成功后接着 /deploy-elevated，成功后自动健康检查）
    AttachLinkSetup.exe /install            仅安装加载项（标准安装）
    AttachLinkSetup.exe /deploy-elevated    仅部署鼎信诺环境（弹 UAC；自动化 / 维护用，界面无此入口）
    AttachLinkSetup.exe /uninstall          一键卸载（用户级 + 机器级看门狗；检测到机器级时弹一次 UAC）
    AttachLinkSetup.exe /uninstall-elevated 仅卸载机器级（弹 UAC，兼容保留）
    AttachLinkSetup.exe /healthcheck        健康检查（纯只读；机器级各项在未部署时显示【跳过】属正常）
    AttachLinkSetup.exe /selfcheck          自检（校验内嵌文件完整性）
    AttachLinkSetup.exe /check-update       检查更新（0=已是最新，20=有新版本，21=检查失败）
    AttachLinkSetup.exe /update             一键更新（0=已交接或已是最新/试运行，22=失败）
      /check-update 与 /update 的附加参数：--update-dry-run（只下载校验不安装）、
      --mirror <加速前缀>、--manifest-url <清单地址>（指向回环地址时允许 http，供本地测试）
- 更新交接参数（由新安装器在交接时自动携带，勿手工构造）：
    --updated-from <旧exe路径> [--wait-pid <旧进程pid>]
    任何启动方式下检测到 --updated-from，都会先尝试把旧安装器文件替换为本安装器的副本
    （等 --wait-pid 指定的旧进程退出后再动文件），替换结果记入日志，然后继续执行本行动作。
    交接模式下，命令行结束时的日志快照写到【旧安装器所在目录】（新安装器运行在 %TEMP% 更新
    临时目录里，快照写在那里用户看不到）；目录不存在/不可写时维持现状写 exe 所在目录。
    界面图形模式请调用本模块的公开入口 handle_update_handoff()（run_cli 已自动调用）。
- 界面层本地演示钩子：环境变量 ATTACHLINK_UPDATE_MANIFEST_URL 提供默认清单地址——只在
  check_update / apply_update 未显式传入 manifest_url 时读取，且只接受回环主机
  （127.0.0.1 / localhost）的 http/https 地址，其余值忽略并写一行日志说明。
- 退出码：0 成功；2 Excel 在运行（安装）/ 卸载预检未通过（Excel 在运行或加载项文件被占用，卸载未执行）；
  3 安装包不完整；8 卸载未完全成功（机器级残留或 UAC 取消）；9 加载项列表清理未完全成功（仅登记残留）；
  10 卸载未完全成功（用户级有残留清单）；12 鼎信诺环境未部署（使用了其他管理员账户授权）；
  1223 UAC 被取消；20 有新版本（/check-update）；21 检查更新失败；22 一键更新失败；
  1 其他失败（含健康检查脚本自身出错）。健康检查的 10 = 有 fail（仅 /healthcheck 与安装后的自动检查）。
  /install-all 任一步失败即返回该步退出码（第一步失败不会继续第二步；第二步失败时用户级已装好，
  结果为「部分完成」而非全失败）。健康检查不改变 /install-all 的退出码，只把结论合并进标题与正文。
"""
import os
import subprocess
import sys
import time
import urllib.parse
from dataclasses import dataclass

import setup_update  # 同目录模块：更新的网络与文件逻辑（清单/下载/校验/解压/交接）都在那里，本模块只做编排与文案

VERSION_FILE = "VERSION.txt"   # 相对 payload_dir()；曾误写成 "payload\VERSION.txt" 导致读到 payload\payload\… 而显示「未知版本」
APP_TITLE = "AttachLink 安装器"
ENV_MANIFEST_URL = "ATTACHLINK_UPDATE_MANIFEST_URL"   # 界面层本地演示钩子：默认清单地址（仅允许回环主机）
_LOOPBACK_HOSTS = {"127.0.0.1", "localhost"}

# 结果类别：成功 / 部分完成（用户级已装好，机器级未完成）/ 失败
OK, PARTIAL, FAIL = "ok", "partial", "fail"
UAC_CANCELLED = 1223

_HINTS = {
    2: "请先关闭所有 Excel（含鼎信诺）后重试",
    3: "安装包不完整，请联系分发人",
    8: "卸载未完全成功（机器级看门狗残留或 UAC 被取消），详见日志",
    9: "加载项列表登记未能完全清理，请关闭所有 Excel 后重新运行卸载；详见日志里的警告",
    10: "卸载未完全成功（有残留），请关闭所有 Excel / WPS 表格后重新运行卸载，仍失败则按清单手动处理",
    12: "请用当前登录账户的管理员身份重新授权；如果当前账户不是管理员，请联系维护者",
    21: "未能获取最新版本信息，可稍后重试，或手动到 https://github.com/jasonisbrave/attachlink/releases 下载",
    22: "自动更新未完成，可重试一次，或手动到 https://github.com/jasonisbrave/attachlink/releases 下载",
}

# 卸载分支 setup.ps1 打印残留清单的标记行（「卸载未完全成功，以下项目未能清除：」），
# 其后以「- 」开头的行即残留项——据此从子进程输出里把清单截进结果正文。
_RESIDUE_MARKER = "以下项目未能清除"


class _TeeQueue:
    """转发队列：把子进程输出行推给界面（或落盘队列）的同时留一份副本，供截取残留清单。"""

    def __init__(self, inner) -> None:
        self.inner = inner
        self.lines = []

    def put(self, x) -> None:
        self.lines.append(x)
        self.inner.put(x)


def _extract_residue_lines(lines) -> list:
    """从子进程输出截取「以下项目未能清除：」之后的「- 残留项」行。"""
    out, on = [], False
    for raw in lines:
        ln = str(raw).rstrip("\r\n")
        if _RESIDUE_MARKER in ln:
            on = True
            continue
        if on:
            s = ln.strip()
            if s.startswith("- "):
                out.append(s[2:].strip())
            elif s:
                break
    return out


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


def save_log_beside_exe(text: str, extra_note: str = "", target_dir: str | None = None) -> str:
    """把日志快照另存到安装器所在目录（AttachLinkSetup.log，覆盖写），返回目标路径。
    target_dir 显式给出时写该目录（交接模式：--updated-from 的旧安装器所在目录——新安装器
    从 %TEMP% 临时目录启动，快照写在那里用户看不到）；目标不可写时静默失败返回空串（提示回落到 %TEMP% 路径）。"""
    target = os.path.join(target_dir if target_dir else exe_dir(), "AttachLinkSetup.log")
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
            "机器级看门狗已部署并请求启动（15 秒后的健康检查会核对其心跳）。\n\n"
            "请完全退出并重新打开鼎信诺和 Excel：在鼎信诺里打开底稿约 2~4 秒后，右键菜单自动出现。")
    if code == UAC_CANCELLED:
        return ActionResult(FAIL, code, "鼎信诺环境未部署（授权被取消）",
                            "管理员授权被取消，鼎信诺环境未部署。\n需要时请重新运行并允许授权。")
    if code == 12:
        return ActionResult(FAIL, code, "鼎信诺环境未部署（使用了其他管理员账户授权）",
                            "安装时使用了另一个管理员账户授权：看门狗计划任务只能登记在该管理员名下，"
                            "当前登录的用户登录时不会启动，鼎信诺环境对当前用户无效。\n\n"
                            "请用当前登录账户的管理员身份重新授权；如果当前账户不是管理员，请联系维护者。",
                            _HINTS[12])
    return _failed(code, f"鼎信诺环境部署失败（退出码 {code}）", "机器级看门狗未能部署。")


def install_all(outq) -> ActionResult:
    """一键安装：先用户级（失败即停），再机器级（UAC 取消 / 失败 = 部分完成），成功后自动健康检查"""
    first = standard_install(outq, "［1/2］安装加载项（用户级）")
    if first.kind != OK:
        return first
    second = deploy_elevated(outq, "［2/2］部署鼎信诺环境（将弹出管理员授权，请点「是」）")
    if second.kind == OK:
        result = ActionResult(
            OK, 0, "一键安装完成",
            "加载项已启用，鼎信诺环境已部署。\n\n"
            "请【完全退出并重新打开 Excel（及鼎信诺）】后使用：\n"
            "· 在单元格上点右键，即可看到「插入附件链接（可多选）」\n"
            "· 在鼎信诺里打开底稿约 2~4 秒后，右键菜单自动出现")
    elif second.code == UAC_CANCELLED:
        result = ActionResult(
            PARTIAL, UAC_CANCELLED, "加载项已安装，鼎信诺环境未部署",
            "管理员授权被取消，鼎信诺环境未部署。\n\n"
            "· 不使用鼎信诺的同事：现在就可以使用，重新打开 Excel 即可。\n"
            "· 需要在鼎信诺里编制底稿的同事：请重新运行并允许授权（或请 IT 用管理员账户协助）。")
    elif second.code == 12:
        result = ActionResult(
            PARTIAL, 12, "加载项已安装，鼎信诺环境未部署（使用了其他管理员账户授权）",
            "安装时使用了另一个管理员账户授权：看门狗计划任务只能登记在该管理员名下，"
            "当前登录的用户登录时不会启动，鼎信诺环境对当前用户无效。\n\n"
            "· 日常使用 Excel 里的加载项功能不受影响，重新打开 Excel 即可。\n"
            "· 需要在鼎信诺里编制底稿的同事：请用当前登录账户的管理员身份重新运行「一键安装」；"
            "如果当前账户不是管理员，请联系维护者。")
    else:
        result = ActionResult(
            PARTIAL, second.code, "加载项已安装，鼎信诺环境部署失败",
            f"机器级看门狗未能部署（退出码 {second.code}）。\n\n"
            "· 日常 Excel 里的加载项功能不受影响，重新打开 Excel 即可使用。\n"
            "· 需要在鼎信诺里编制底稿的同事：请把日志反馈给分发人。")
    # 安装后自动健康检查：等 15 秒心跳让看门狗完成首轮记录；自身出错只记一行，不影响安装结论与退出码
    try:
        hc = health_check(outq, heartbeat_wait_s=15)
        if hc.kind == FAIL:
            # hc.title 形如「健康检查：N 项异常」→ 取「N 项异常」
            tail = hc.title.split("：", 1)[1] if "：" in hc.title else hc.title
            if result.kind == OK:
                result.kind = PARTIAL
                result.title = f"安装完成，但健康检查发现 {tail}"
            result.body += f"\n■ 健康检查：发现 {tail}（详情见日志，可点「健康检查」重新查看）"
        elif hc.kind == PARTIAL:
            # 只有「需留意」（无异常）：kind 与退出码不变（README：健康检查不改变 /install-all 的退出码），
            # 但结论必须并进正文——README 承诺「有需留意 / 异常项时结果会列出」
            tail = hc.title.split("：", 1)[1] if "：" in hc.title else hc.title
            detail = "\n".join(hc.body.splitlines()[1:])   # hc.body 首行是汇总，其后逐条为「· id 消息（处置：提示）」
            result.body += f"\n■ 健康检查：{tail}" + (f"：\n{detail}" if detail else "（详情见日志）")
    except Exception as e:  # noqa: BLE001
        outq.put(f"\n[健康检查] 自身出错（不影响安装结果）：{e}\n")
    return result


def uninstall(outq) -> ActionResult:
    """一键卸载：用户级 + 机器级看门狗（setup.ps1 -Uninstall 自动检测并弹一次 UAC）"""
    outq.put("\n一键卸载（用户级加载项 + 鼎信诺看门狗）…\n")
    tee = _TeeQueue(outq)
    code = Runner(tee).run(_ps("scripts\\setup.ps1", ["-Uninstall", "-NonInteractive"]), payload_dir())
    if code == 0:
        _uninstall_cleanup_tmp()
        return ActionResult(
            OK, 0, "卸载完成",
            "用户级加载项与机器级看门狗均已卸载（按事实核验无残留）。\n\n"
            "已打开的 Excel 会话里，右键菜单会保留到该会话关闭；重新打开 Excel 后不再出现。")
    if code == 2:
        return ActionResult(FAIL, 2, "卸载未执行：Excel 正在运行或加载项文件被占用",
                            "检测到 Excel 正在运行，或加载项文件被其它程序（含 WPS 表格）占用——"
                            "为避免「删不掉文件、旧会话退出时又把加载项登记写回注册表」的假卸载，本次卸载没有执行。",
                            _HINTS[2])
    if code == 10:
        residue = _extract_residue_lines(tee.lines)
        detail = ("\n".join("· " + r for r in residue)) if residue else "详见日志里的残留清单。"
        return _failed(code, "卸载未完全成功（有残留）",
                       "以下项目未能清除。请关闭所有 Excel / WPS 表格后重新运行卸载；仍失败则按清单手动处理：\n" + detail)
    if code == 9:
        return _failed(code, "卸载未完全成功（加载项列表清理未完全成功）",
                       "加载项已卸载，但加载项列表的登记未能完全清理。")
    if code == 8:
        return _failed(code, "卸载未完全成功", "卸载已执行，但机器级看门狗仍有残留，或管理员授权被取消。")
    return _failed(code, f"卸载失败（退出码 {code}）", "卸载未能完成。")


def _uninstall_cleanup_tmp() -> None:
    """卸载收尾（两级都卸干净后）：删除看门狗日志（存在才删），并清理更新临时根下超过 24 小时的旧子目录。
    不碰正在运行的子目录——卸载器自身可能正从某个 GUID 子目录里运行（_clean_stale_tmp_dirs 只删超时目录）。"""
    try:
        wd_log = os.path.join(os.environ.get("TEMP", "."), "AttachLinkWatchdog.log")
        if os.path.isfile(wd_log):
            os.remove(wd_log)
    except OSError:
        pass
    try:
        tmp_root = os.path.join(os.environ.get("TEMP", "."), setup_update._TMP_DIR_NAME)
        if os.path.isdir(tmp_root):
            setup_update._clean_stale_tmp_dirs(tmp_root)
    except OSError:
        pass


def uninstall_elevated_only(outq) -> ActionResult:
    """兼容保留：仅卸载机器级"""
    outq.put("\n仅卸载机器级看门狗…\n")
    code = Runner(outq).run_elevated("build\\install-elevated.ps1 -Uninstall")
    if code == 0:
        return ActionResult(OK, 0, "机器级看门狗已卸载", "机器级看门狗已卸载（按事实核验无残留）。用户级加载项未改动。")
    if code == 8:
        return _failed(code, "机器级卸载未完全成功", "机器级卸载已执行，但仍有残留（计划任务 / 部署目录 / HKLM 信任位置 / 看门狗进程），详见日志里的残留清单。")
    return _failed(code, f"机器级卸载失败（退出码 {code}）", "机器级看门狗未能卸载。")


# ---------------- 健康检查（纯只读） ----------------

# 界面按这五个前缀着色，不要改文案
_HC_PREFIX = {
    "ok": "【通过】", "warn": "【留意】", "fail": "【异常】", "skip": "【跳过】", "info": "【信息】",
}
_HC_STATUSES = frozenset(_HC_PREFIX)


class _ListQueue:
    """健康检查用的捕获队列：子进程原始协议行先进这里，解析后再推人读行进界面日志"""

    def __init__(self) -> None:
        self.lines = []

    def put(self, x) -> None:
        self.lines.append(x)


def _parse_hc_lines(lines):
    """解析 healthcheck.ps1 的协议行 → (items, summary)。
    items：[{'id','status','message','hint'}]（按输出顺序；HCHINT 回填到同 id 条目）；
    summary：{'ok','warn','fail','skip','info'} 计数字典或 None。非协议行忽略。"""
    items = []
    hints = {}
    summary = None
    for raw in lines:
        line = str(raw).lstrip("\ufeff").rstrip("\r\n")
        if not line.strip():
            continue
        parts = line.split("\t")
        if parts[0] == "HC" and len(parts) >= 3 and parts[1] == "SUMMARY":
            summary = {}
            for tok in " ".join(parts[2:]).split():
                key, _, val = tok.partition("=")
                summary[key] = int(val) if val.isdigit() else 0
            continue
        if parts[0] == "HC" and len(parts) == 4 and parts[2] in _HC_STATUSES:
            items.append({"id": parts[1], "status": parts[2], "message": parts[3], "hint": ""})
            continue
        if parts[0] == "HCHINT" and len(parts) == 3:
            hints[parts[1]] = parts[2]
            continue
    for it in items:
        it["hint"] = hints.get(it["id"], "")
    return items, summary


def health_check(outq, heartbeat_wait_s: int = 0) -> ActionResult:
    """健康检查：运行 scripts\\healthcheck.ps1（纯只读，不写注册表 / 不改文件 / 不动计划任务），
    把人读得懂的行逐条推入 outq（前缀固定为【通过】【留意】【异常】【跳过】【信息】，处置提示以
    「↳」跟随同前缀推送）。kind：有 fail → FAIL；只有 warn → PARTIAL；否则 OK；code 透传脚本退出码
    （0 / 10 / 1）。机器级整组为【跳过】（未部署鼎信诺环境）时不计入 warn / fail。"""
    outq.put("\n健康检查（只读：不写注册表、不改文件、不动计划任务；未部署机器级时相关项显示【跳过】属正常）…\n")
    cap = _ListQueue()
    args = _ps("scripts\\healthcheck.ps1",
               ["-PayloadDir", payload_dir(), "-HeartbeatWaitSeconds", str(int(heartbeat_wait_s))])
    code = Runner(cap).run(args, payload_dir())
    items, summary = _parse_hc_lines(cap.lines)
    if not items or summary is None:
        # 脚本没跑起来 / 输出不可解析：原始输出透传进日志便于排查
        outq.put("【异常】健康检查脚本未产生可解析的输出，原始输出如下：")
        for ln in cap.lines:
            outq.put("  " + str(ln).rstrip())
        return ActionResult(FAIL, code if code else 1, "健康检查：未获得结果",
                            "健康检查脚本没有产生可解析的输出。", None)
    for it in items:
        outq.put(f"{_HC_PREFIX[it['status']]}{it['id']} {it['message']}")
        if it["hint"]:
            outq.put(f"{_HC_PREFIX[it['status']]}  ↳ {it['hint']}")
    n_fail = summary.get("fail", 0)
    n_warn = summary.get("warn", 0)
    fails = [it for it in items if it["status"] == "fail"]
    body_lines = [f"通过 {summary.get('ok', 0)} 项，需留意 {n_warn} 项，异常 {n_fail} 项，"
                  f"跳过 {summary.get('skip', 0)} 项（另有信息 {summary.get('info', 0)} 条；"
                  "跳过=未部署机器级看门狗时属正常）"]
    for it in items:
        if it["status"] in ("warn", "fail"):
            msg = f"· {it['id']} {it['message']}"
            if it["hint"]:
                msg += f"（处置：{it['hint']}）"
            body_lines.append(msg)
    if fails:
        kind, title = FAIL, f"健康检查：{n_fail} 项异常"
        hint = fails[0]["hint"] or None
    elif n_warn:
        kind, title = PARTIAL, f"健康检查：{n_warn} 项需留意"
        hint = None
    else:
        kind, title = OK, "健康检查：全部正常"
        hint = None
    outq.put(f"[健康检查] {title}；退出码 {code}")
    return ActionResult(kind, code, title, "\n".join(body_lines), hint)



# ---------------- 检查更新 / 一键更新 ----------------

def _env_manifest_url(outq=None):
    """界面层本地测试钩子：清单地址未显式传入时读环境变量 ATTACHLINK_UPDATE_MANIFEST_URL。
    只接受回环主机（127.0.0.1 / localhost）的 http/https 地址；其余值忽略并写一行日志说明，返回 None。"""
    url = os.environ.get(ENV_MANIFEST_URL, "").strip()
    if not url:
        return None
    parts = urllib.parse.urlsplit(url)
    if parts.scheme in ("http", "https") and (parts.hostname or "").lower() in _LOOPBACK_HOSTS:
        return url
    line = f"已忽略环境变量 {ENV_MANIFEST_URL}（仅允许回环地址 127.0.0.1 / localhost 的 http/https）：{url}"
    if outq is not None:
        outq.put(line + "\n")
    else:
        append_log_file(line)
    return None


def check_update(outq, mirror=None, manifest_url=None) -> ActionResult:
    """检查更新：拉发布清单并与当前版本比较（失败转成失败结果，不抛异常）。
    manifest_url 未显式传入时读环境变量 ATTACHLINK_UPDATE_MANIFEST_URL（仅回环地址，界面层本地演示用）。"""
    outq.put("\n检查更新…\n")
    current = read_version()
    if manifest_url is None:
        manifest_url = _env_manifest_url(outq)
    try:
        manifest, worked = setup_update.fetch_manifest(
            mirror=mirror, manifest_url=manifest_url, current_version=current,
            log=lambda s: outq.put(s + "\n"))
    except setup_update.UpdateError as e:
        outq.put(f"\n{e.message}\n")
        return ActionResult(
            FAIL, 21, "检查更新失败",
            "未能获取最新版本信息。",
            _HINTS[21])
    latest = manifest.get("version", "?")
    outq.put(f"检查更新：最新 v{latest}（当前 v{current}）\n")
    if setup_update.is_newer(latest, current):
        # 发布说明地址（清单有才带）：界面据 body 里「发布说明：」一行显示链接
        notes = manifest.get("notes_url")
        notes_line = f"\n发布说明：{notes}" if notes else ""
        if notes:
            outq.put(f"发布说明：{notes}\n")
        return ActionResult(
            PARTIAL, 20, f"发现新版本 v{latest}",
            f"最新版本为 v{latest}（当前 v{current}）。{notes_line}",
            "点「一键更新」即可自动下载并安装")
    mirror_note = ""
    if worked:
        # 经加速镜像取得（直连不可达/未用）且镜像说没有新版：注明可能有缓存延迟（B4 复核只在「说没有新版」时发生）
        mirror_note = ("（据加速镜像，新版本发布后可能有缓存延迟；如刚发布新版，可稍后再查或到 Releases 页面手动下载）")
    return ActionResult(OK, 0, "已是最新版本", f"当前版本 v{current} 已是最新。{mirror_note}")


def apply_update(outq, action="/install-all") -> ActionResult:
    """一键更新：下载并校验新安装器后交接给新安装器安装（界面入口；成功后 hint=EXIT_AFTER，界面据此关闭自己）"""
    return _apply_update(outq, action=action)


def _apply_update(outq, action="/install-all", dry_run=False, mirror=None, manifest_url=None) -> ActionResult:
    """apply_update 的实现体（/update 命令行会额外传 dry_run / mirror / manifest_url）。
    manifest_url 未显式传入时读环境变量 ATTACHLINK_UPDATE_MANIFEST_URL（仅回环地址，界面层本地演示用）。"""
    if manifest_url is None:
        manifest_url = _env_manifest_url(outq)

    def _out(line):
        outq.put(str(line) + "\n")

    def _progress(done, total):
        pct = f"{done * 100 // total}" if total and total > 0 else "-"
        outq.put(f"【进度】{done}\t{total}\t{pct}%\n")

    summary = setup_update.run_update(dry_run=dry_run, mirror=mirror, manifest_url=manifest_url,
                                      action=action, out=_out, progress=_progress)
    outq.put("\n" + summary + "\n")
    if summary.startswith("已启动新版本安装器"):
        return ActionResult(
            OK, 0, "已启动新版本安装器",
            "本窗口即将关闭；新版本会自动开始安装，如弹出管理员授权请点「是」。",
            "EXIT_AFTER")
    if summary.startswith("当前已是最新版本"):
        return ActionResult(OK, 0, "已是最新版本", summary)
    if summary.startswith("试运行完成"):
        return ActionResult(OK, 0, "试运行完成", summary)
    return ActionResult(FAIL, 22, "更新失败", summary, "")


# ---------------- 静默模式 ----------------

class _FileQueue:
    """静默模式用的队列替身：日志行直接落盘"""

    def put(self, line: str) -> None:
        append_log_file(line.rstrip("\n"))


def selfcheck() -> int:
    pdir = payload_dir()
    need = ["VERSION.txt", "dist\\AttachLink.xlam", "scripts\\setup.ps1", "scripts\\watchdog.ps1",
            "dist\\AttachLink.WatchdogLauncher.exe", "scripts\\healthcheck.ps1", "build\\install.ps1",
            "build\\install-elevated.ps1", "run-elevated.ps1", "elevated-capture.ps1"]
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
    "/healthcheck": health_check,
}


def handle_update_handoff() -> dict:
    """（公开入口，界面图形模式同样调用）更新交接自检：命令行带 --updated-from <旧exe> [--wait-pid <pid>] 时，
    先等旧进程退出并尝试把旧安装器文件替换为本安装器的副本，结果写进日志，再继续原本的动作。
    返回结构化结果 {"attempted", "ok", "message", "old_exe", "status", "backup"}：未尝试交接
    attempted=False、ok=None；尝试后 ok 取替换结果（异常视同失败）——调用方必须据 ok / status
    如实呈现「旧安装器文件是否已更新」，替换失败 / 旧文件缺失时旧 exe 仍是旧版或不存在，
    不得无条件告知用户已更新；status ∈ replaced / recreated / same / missing / failed，
    backup 为回滚失败时保留的备份路径（其余情况为 None）。
    run_cli 已自动调用；界面侧在创建窗口前调用一次即可。"""
    hand = {"attempted": False, "ok": None, "message": "", "old_exe": "", "status": None, "backup": None}
    argv = sys.argv
    old_exe = None
    wait_pid = None
    for i, token in enumerate(argv):
        low = token.lower()
        if low == "--updated-from" and i + 1 < len(argv):
            old_exe = argv[i + 1]
        elif low == "--wait-pid" and i + 1 < len(argv):
            wait_pid = argv[i + 1]
    if not old_exe:
        return hand
    hand["attempted"] = True
    hand["old_exe"] = old_exe
    append_log_file(f"[交接] 检测到 --updated-from：旧安装器={old_exe} 等待进程={wait_pid or '无'}")
    try:
        res = setup_update.replace_old_exe(old_exe, wait_pid=wait_pid)
        ok = bool(res.get("ok"))
        message = str(res.get("message", ""))
        hand["status"] = res.get("status")
        hand["backup"] = res.get("backup")
    except Exception as e:  # noqa: BLE001  交接失败不阻断本行动作
        ok, message = False, f"替换旧安装器时出现异常：{e}"
    hand["ok"] = ok
    hand["message"] = message
    append_log_file(f"[交接] {'完成' if ok else '失败'}：{message}")
    return hand


def handoff_summary_line(hand: dict | None, new_exe_dir: str) -> str:
    """交接结果给用户看的一行说明（纯函数，便于单测）。
    未尝试 → 空串；按 status 区分措辞：replaced →「旧安装器文件已更新」；recreated → 原位置此前
    已不存在、已按新版本重建；same → 旧安装器即本安装器；missing → 旧文件不存在且原位置不可写，
    如实说明并给出本安装器位置；failed → 如实说明未能替换并给出手动复制建议（等待旧进程超时 /
    杀毒锁定 / 目录不可写时旧 exe 仍是旧版），回滚失败保留备份时追加备份路径与改回提示。"""
    if not hand or not hand.get("attempted"):
        return ""
    status = hand.get("status")
    if hand.get("ok"):
        if status == "recreated":
            return "已由旧版本安装器交接启动；原位置的安装器此前已不存在，已按新版本重建"
        if status == "same":
            return "已由旧版本安装器交接启动（旧安装器即本安装器）"
        # replaced（或旧调用方未提供 status 的兼容回退）
        return "已由旧版本安装器交接启动，旧安装器文件已更新"
    if status == "missing":
        return (f"已由旧版本安装器交接启动；{hand.get('message', '')}"
                f"——本安装器位于 {new_exe_dir}，请把它放到固定位置")
    line = (f"已由旧版本安装器交接启动；旧安装器文件未能替换（{hand.get('message', '')}）"
            f"——本安装器位于 {new_exe_dir}，建议把它复制到原来的安装器位置")
    if hand.get("backup"):
        line += f"；旧文件的备份保留在：{hand['backup']}——原位置目前没有安装器，请把该备份改回原名"
    return line


# 兼容别名：上一轮实现名为 _handle_updated_from，run_cli 与既有测试两者都可用
_handle_updated_from = handle_update_handoff


def _log_beside_dir(argv=None):
    """交接模式下的日志快照目录：命令行带 --updated-from <旧exe> 时返回旧安装器所在目录
    （新安装器运行在 %TEMP% 更新临时目录里，快照写在那里用户看不到）；无交接参数、
    或旧目录不存在时返回 None（维持现状：写到 exe 所在目录）。纯函数，便于单测。"""
    argv = sys.argv if argv is None else argv
    for i, token in enumerate(argv):
        if token.lower() == "--updated-from" and i + 1 < len(argv):
            d = os.path.dirname(os.path.abspath(argv[i + 1]))
            return d if os.path.isdir(d) else None
    return None


def _cli_extras() -> dict:
    """解析 /check-update 与 /update 的附加参数（--update-dry-run / --mirror / --manifest-url）"""
    extras = {"mirror": None, "manifest_url": None, "dry_run": False, "action": "/install-all"}
    argv = sys.argv[1:]
    i = 0
    while i < len(argv):
        low = argv[i].lower()
        if low == "--mirror" and i + 1 < len(argv):
            extras["mirror"] = argv[i + 1]
            i += 2
        elif low == "--manifest-url" and i + 1 < len(argv):
            extras["manifest_url"] = argv[i + 1]
            i += 2
        elif low == "--update-dry-run":
            extras["dry_run"] = True
            i += 1
        else:
            i += 1
    return extras


def run_cli(mode: str) -> int:
    logf = log_file_path()
    append_log_file(f"===== AttachLinkSetup {mode} @ {time.strftime('%Y-%m-%d %H:%M:%S')} =====")
    hand = handle_update_handoff()
    if mode == "/selfcheck":
        return selfcheck()
    extras = _cli_extras()
    if mode == "/check-update":
        result = safe(lambda q: check_update(q, extras["mirror"], extras["manifest_url"]), _FileQueue())
    elif mode == "/update":
        result = safe(lambda q: _apply_update(q, action=extras["action"], dry_run=extras["dry_run"],
                                              mirror=extras["mirror"], manifest_url=extras["manifest_url"]),
                      _FileQueue())
    else:
        action = CLI_ACTIONS.get(mode)
        if action is None:
            print(f"未知参数：{mode}")
            return 1
        result = safe(action, _FileQueue())
    # 交接失败不改变退出码（安装动作本身可能已成功），但结果行必须如实带提示：
    # 此时旧 exe 仍是旧版，用户需要知道更新并未落到原安装器位置
    note = ""
    if hand.get("attempted") and hand.get("ok") is False:
        note = f"；注意：旧安装器文件未替换：{hand['message']}"
    append_log_file(f"[结果] {result.title}（退出码 {result.code}）{note}")
    print(f"[AttachLinkSetup] {result.title}；退出码 {result.code}；完整日志：{logf}{note}")
    # 日志快照另存（交接模式写到旧安装器所在目录——新安装器运行在 %TEMP% 更新临时目录里，
    # 快照写在那里用户看不到；无交接参数或目录不可用则维持现状写 exe 所在目录）
    beside = save_log_beside_exe(full_log_text(), f"[silent {mode} exit={result.code}] {result.title}",
                                 target_dir=_log_beside_dir())
    if beside:
        print(f"[AttachLinkSetup] 日志快照已另存：{beside}")
    return result.code

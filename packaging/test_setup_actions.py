# setup_actions.py 无头单测：不启动 PowerShell，用假 Runner 模拟各步退出码，核对编排与结果文案
import contextlib
import importlib.util
import io
import os
import tempfile

# AttachLink 安装器动作层（setup_actions.py）无头单测：不启动 PowerShell，用假 Runner 模拟各步退出码，核对编排与结果文案。
# 用法：python packaging\test_setup_actions.py   期望末行 RESULT: PASS
# 本文件不进安装包：PyInstaller 只打包被 setup_gui.py 导入的模块。

spec = importlib.util.spec_from_file_location("setup_actions", os.path.join(os.path.dirname(os.path.abspath(__file__)), "setup_actions.py"))
sa = importlib.util.module_from_spec(spec); spec.loader.exec_module(sa)

class Q:
    def __init__(self): self.lines = []
    def put(self, x): self.lines.append(x)

ok = True
def check(cond, label):
    global ok
    print(("pass: " if cond else "FAIL: ") + label)
    if not cond: ok = False

def with_codes(codes):
    """让 Runner.run / run_elevated 依次返回 codes（Standard 用户级在前，Elevated 机器级在后）"""
    seq = {"run": iter([c for k, c in codes if k == "S"]), "run_elevated": iter([c for k, c in codes if k == "E"])}
    class FakeRunner:
        def __init__(self, outq): pass
        def run(self, args, cwd): return next(seq["run"])
        def run_elevated(self, target): return next(seq["run_elevated"])
    return FakeRunner

def run(action, codes, label=""):
    q = Q(); real = sa.Runner; sa.Runner = with_codes(codes)
    try: return sa.safe(action, q), q
    finally: sa.Runner = real

r, q = run(sa.install_all, [("S", 0), ("E", 0)])
check(r.kind == sa.OK and r.code == 0, "一键安装 两步成功 → OK / 0")
check("［1/2］" in "".join(map(str, q.lines)) and "［2/2］" in "".join(map(str, q.lines)), "一键安装 日志带 [1/2] [2/2] 步骤标题")
check("重新打开 Excel（及鼎信诺）" in r.dialog_text("X"), "一键安装成功文案")

r, q = run(sa.install_all, [("S", 2)])
check(r.kind == sa.FAIL and r.code == 2 and r.hint == "请先关闭所有 Excel（含鼎信诺）后重试", "第一步 Excel 在运行 → FAIL / 2，带原因提示")
check(not any("［2/2］" in str(x) for x in q.lines), "第一步失败不继续第二步")

r, q = run(sa.install_all, [("S", 0), ("E", 1223)])
check(r.kind == sa.PARTIAL and r.code == 1223, "第二步 UAC 取消 → PARTIAL / 1223")
check("不使用鼎信诺的同事" in r.dialog_text("日志位置"), "UAC 取消文案区分两类同事")

r, q = run(sa.install_all, [("S", 0), ("E", 5)])
check(r.kind == sa.PARTIAL and r.code == 5, "第二步其他失败 → PARTIAL / 5")

r, q = run(sa.standard_install, [("S", 0)])
check("一键安装" in r.dialog_text("X") and "更多选项" not in r.dialog_text("X"), "仅安装加载项成功文案：指向「一键安装」，且不再提单独部署")

r, q = run(sa.deploy_elevated, [("E", 1223)])
check(r.kind == sa.FAIL and r.code == 1223 and "授权被取消" in r.title, "单独部署 UAC 取消 → FAIL / 1223")

r, q = run(sa.uninstall, [("S", 8)])
check(r.kind == sa.FAIL and r.code == 8, "卸载 exit 8 → FAIL / 8")
r, q = run(sa.uninstall, [("S", 9)])
check(r.kind == sa.FAIL and r.code == 9 and "加载项列表" in r.title, "卸载 exit 9 → FAIL / 9")
r, q = run(sa.uninstall, [("S", 0)])
check(r.kind == sa.OK, "卸载成功 → OK")

def boom(outq): raise RuntimeError("模拟异常")
r = sa.safe(boom, Q())
check(r.kind == sa.FAIL and r.code == 1 and "模拟异常" in r.body, "safe() 把异常转成 FAIL / 1")

check(set(sa.CLI_ACTIONS) == {"/install-all", "/install", "/deploy-elevated", "/uninstall", "/uninstall-elevated", "/healthcheck"}, "静默参数表完整（含 /healthcheck）")

# 版本读取（回归：VERSION_FILE 曾被重复拼接成 payload\payload\VERSION.txt，发布包显示「未知版本」）
with tempfile.TemporaryDirectory() as d:
    os.makedirs(os.path.join(d, "payload"))
    with open(os.path.join(d, "payload", "VERSION.txt"), "w", encoding="utf-8") as fh:
        fh.write("9.9.9\n")
    real_payload_dir = sa.payload_dir
    sa.payload_dir = lambda: os.path.join(d, "payload")
    try:
        check(sa.read_version() == "9.9.9", "read_version：在 payload_dir() 下直接读 VERSION.txt")
        with contextlib.redirect_stdout(io.StringIO()):
            check(sa.selfcheck() == 3, "selfcheck：载荷齐全但缺 VERSION.txt 之外的文件时返回 3（临时目录无载荷）")
            for n in ["dist\\AttachLink.xlam", "scripts\\setup.ps1", "scripts\\watchdog.ps1", "dist\\AttachLink.WatchdogLauncher.exe",
                      "build\\install.ps1", "build\\install-elevated.ps1", "run-elevated.ps1", "elevated-capture.ps1"]:
                os.makedirs(os.path.dirname(os.path.join(d, "payload", n)), exist_ok=True)
                open(os.path.join(d, "payload", n), "w").close()
            with contextlib.redirect_stdout(io.StringIO()):
                check(sa.selfcheck() == 3, "selfcheck：缺第 10 个文件 scripts\\healthcheck.ps1 → 3")
            open(os.path.join(d, "payload", "scripts", "healthcheck.ps1"), "w").close()
            with contextlib.redirect_stdout(io.StringIO()):
                check(sa.selfcheck() == 0, "selfcheck：10 个载荷齐全且版本可读 → 0")
            os.remove(os.path.join(d, "payload", "dist", "AttachLink.WatchdogLauncher.exe"))
            buf_miss = io.StringIO()
            with contextlib.redirect_stdout(buf_miss):
                check(sa.selfcheck() == 3, "selfcheck：缺 dist\\AttachLink.WatchdogLauncher.exe → 3")
            check("缺失" in buf_miss.getvalue() and "dist\\AttachLink.WatchdogLauncher.exe" in buf_miss.getvalue(),
                  "selfcheck：缺启动器 exe 时缺失清单报出该文件名")
            open(os.path.join(d, "payload", "dist", "AttachLink.WatchdogLauncher.exe"), "w").close()
            open(os.path.join(d, "payload", "VERSION.txt"), "w").close()   # 版本文件变空
            with contextlib.redirect_stdout(io.StringIO()):
                check(sa.selfcheck() == 3, "selfcheck：VERSION.txt 为空 → 3（不再静默显示未知版本）")
    finally:
        sa.payload_dir = real_payload_dir

# run_cli 的纯逻辑分支（不触发安装）：未知参数 → 1
with contextlib.redirect_stdout(io.StringIO()):
    check(sa.run_cli("/no-such-mode") == 1, "未知静默参数 → 1")

# =====================================================================
# 健康检查（v1.5.0）：协议行解析、kind 映射、install_all 自动合并、CLI 透传（全部打桩，不跑真脚本）
# =====================================================================

HC_FAIL_LINES = [
    "HC\tuser.file\tok\t用户级加载项文件在位（C:\\Users\\u\\AppData\\Roaming\\Microsoft\\AddIns\\AttachLink.xlam）",
    "HC\tmachine.deploy\tok\t机器级看门狗已部署（计划任务 + Program Files 目录均在位）",
    "HC\tmachine.heartbeat\tfail\t看门狗可能未运行：日志 300 秒未刷新",
    "HCHINT\tmachine.heartbeat\t确认没有安全软件禁用任务 → 管理员 PowerShell 执行 Start-ScheduledTask",
    "HC\tmachine.conf\tinfo\t父进程白名单生效中：sjp7.exe（仅附着这些进程拉起的 Excel）",
    "HC\tSUMMARY\tok=2 warn=0 fail=1 skip=0 info=1",
]
HC_OK_LINES = [
    "HC\tuser.file\tok\t用户级加载项文件在位",
    "HC\tmachine.deploy\tskip\t未部署机器级看门狗——不使用鼎信诺（管理员运行）时属正常",
    "HC\tmachine.heartbeat\tskip\t未部署机器级看门狗——不使用鼎信诺（管理员运行）时属正常",
    "HC\tmachine.task\tskip\t未部署机器级看门狗——不使用鼎信诺（管理员运行）时属正常",
    "HC\tSUMMARY\tok=1 warn=0 fail=0 skip=3 info=0",
]
HC_WARN_LINES = [
    "HC\tuser.version\twarn\t用户级副本与安装包版本不一致（前 16 位 abc 与 def），请重新安装",
    "HCHINT\tuser.version\t点「一键安装」覆盖安装即可同步版本",
    "HC\tSUMMARY\tok=0 warn=1 fail=0 skip=0 info=0",
]


def with_install_and_hc(step_codes, hc_lines=None, hc_code=0):
    """打桩 Runner：healthcheck.ps1 的调用推送 hc_lines 并返回 hc_code；其余 run / run_elevated 依次返回 step_codes。
    返回 (FakeRunner 类, hc_args 字典)——hc_args["args"] 记录了健康检查收到的命令行参数。"""
    steps = iter(step_codes)
    hc_args = {}

    class FakeRunner:
        def __init__(self, outq): self.outq = outq
        def run(self, args, cwd):
            if any("healthcheck.ps1" in str(a) for a in args):
                hc_args["args"] = [str(a) for a in args]
                for ln in (hc_lines if hc_lines is not None else []):
                    self.outq.put(ln)
                return hc_code
            return next(steps)
        def run_elevated(self, target): return next(steps)
    return FakeRunner, hc_args


def run_patched(action, fake_runner):
    q = Q(); real = sa.Runner; sa.Runner = fake_runner
    try: return sa.safe(action, q), q
    finally: sa.Runner = real

# --- 协议行解析：条目提取 + 状态计数 + HCHINT 回填 + 非协议行忽略 + BOM 容忍 ---
items, summary = sa._parse_hc_lines(HC_FAIL_LINES)
check(len(items) == 4 and summary == {"ok": 2, "warn": 0, "fail": 1, "skip": 0, "info": 1}, "协议行解析：条目数与状态计数")
check(items[2]["id"] == "machine.heartbeat" and items[2]["hint"].startswith("确认没有安全软件"), "协议行解析：HCHINT 回填到同 id 条目")
items2, summary2 = sa._parse_hc_lines(["\ufeffHC\tuser.file\tok\t首行带 BOM", "垃圾行", "", "$ healthcheck ..."])
check(len(items2) == 1 and items2[0]["status"] == "ok" and summary2 is None, "协议行解析：BOM 行可解析、非协议行忽略、无 SUMMARY → None")

# --- health_check 三种 kind 映射 ---
fake, _ = with_install_and_hc([], HC_FAIL_LINES, 10)
r, q = run_patched(sa.health_check, fake)
check(r.kind == sa.FAIL and r.code == 10 and r.title == "健康检查：1 项异常", "health_check：有 fail → FAIL / 10 /「健康检查：1 项异常」")
check(r.hint == "确认没有安全软件禁用任务 → 管理员 PowerShell 执行 Start-ScheduledTask", "health_check：hint=第一个 fail 项的处置提示")
check(r.body.splitlines()[0].startswith("通过 2 项，需留意 0 项，异常 1 项，跳过 0 项"), "health_check：body 首行是汇总")
check(any(l.startswith("· machine.heartbeat") and "Start-ScheduledTask" in l for l in r.body.splitlines()), "health_check：body 每个 fail 一行（消息+处置提示）")
joined = "".join(map(str, q.lines))
check("【通过】user.file" in joined and "【异常】machine.heartbeat" in joined and "【信息】machine.conf" in joined, "health_check：人读行带【通过】【异常】【信息】前缀（info 也推）")
check("【异常】  ↳ 确认没有安全软件" in joined, "health_check：处置提示以「↳」跟随同前缀推送")
check("[健康检查] 健康检查：1 项异常；退出码 10" in joined, "health_check：末尾记一条结论+退出码")

fake, _ = with_install_and_hc([], HC_WARN_LINES, 0)
r, _ = run_patched(sa.health_check, fake)
check(r.kind == sa.PARTIAL and r.code == 0 and r.title == "健康检查：1 项需留意", "health_check：只有 warn → PARTIAL / 0 /「健康检查：1 项需留意」")

fake, _ = with_install_and_hc([], HC_OK_LINES, 0)
r, _ = run_patched(sa.health_check, fake)
check(r.kind == sa.OK and r.code == 0 and r.title == "健康检查：全部正常", "health_check：机器级整组 skip 不计入 warn/fail → OK / 全部正常")
check(r.hint is None and "需留意 0 项，异常 0 项" in r.body.splitlines()[0], "health_check：全过时 hint=None、汇总计数正确")

fake, _ = with_install_and_hc([], ["HC\tscript\tfail\t健康检查脚本自身出错：boom", "HC\tSUMMARY\tok=0 warn=0 fail=1 skip=0 info=0"], 1)
r, _ = run_patched(sa.health_check, fake)
check(r.kind == sa.FAIL and r.code == 1 and "boom" in r.body, "health_check：脚本自身出错 → FAIL / 1（非致命，只记一行）")

fake, hc_args = with_install_and_hc([], HC_OK_LINES, 0)
r, _ = run_patched(lambda outq: sa.health_check(outq, heartbeat_wait_s=7), fake)
check("-HeartbeatWaitSeconds" in hc_args["args"] and hc_args["args"][hc_args["args"].index("-HeartbeatWaitSeconds") + 1] == "7"
      and "-PayloadDir" in hc_args["args"], "health_check：命令行带 -PayloadDir 与 -HeartbeatWaitSeconds（透传参数）")

fake, _ = with_install_and_hc([], ["$ healthcheck.ps1 ...", "输出完全不像协议"], 1)
r, _ = run_patched(sa.health_check, fake)
check(r.kind == sa.FAIL and r.code == 1 and r.title == "健康检查：未获得结果", "health_check：输出不可解析 → FAIL / 1 / 未获得结果")

# --- install_all：第二步之后自动健康检查并合并 ---
fake, hc_args = with_install_and_hc([0, 0], HC_FAIL_LINES, 10)
r, q = run_patched(sa.install_all, fake)
check(r.kind == sa.PARTIAL and r.code == 0, "install_all：两步成功+健康 fail → OK 降级 PARTIAL，退出码不变（0）")
check(r.title == "安装完成，但健康检查发现 1 项异常", "install_all：降级标题「安装完成，但健康检查发现 N 项异常」")
check("■ 健康检查：发现 1 项异常（详情见日志，可点「健康检查」重新查看）" in r.body, "install_all：body 末尾追加 ■ 健康检查行")
check(hc_args["args"][hc_args["args"].index("-HeartbeatWaitSeconds") + 1] == "15", "install_all：自动健康检查等待 15 秒心跳")
check("【异常】machine.heartbeat" in "".join(map(str, q.lines)), "install_all：健康检查人读行进界面日志")

fake, _ = with_install_and_hc([0, 1223], HC_FAIL_LINES, 10)
r, _ = run_patched(sa.install_all, fake)
check(r.kind == sa.PARTIAL and r.code == 1223 and r.title == "加载项已安装，鼎信诺环境未部署", "install_all：UAC 取消+健康 fail → 仍 PARTIAL / 1223，标题保持")
check("■ 健康检查：发现 1 项异常" in r.body, "install_all：PARTIAL 的 body 也追加健康检查行")

fake, hc_args = with_install_and_hc([2], HC_FAIL_LINES, 10)
r, _ = run_patched(sa.install_all, fake)
check(r.kind == sa.FAIL and r.code == 2 and "args" not in hc_args, "install_all：第一步失败 → 直接返回，不跑健康检查")

fake, _ = with_install_and_hc([0, 0], HC_OK_LINES, 0)
r, _ = run_patched(sa.install_all, fake)
check(r.kind == sa.OK and r.code == 0 and r.title == "一键安装完成" and "■ 健康检查" not in r.body, "install_all：健康全过（含机器级 skip）→ 保持 OK，body 不追加")

fake, _ = with_install_and_hc([0, 0], HC_WARN_LINES, 0)
r, _ = run_patched(sa.install_all, fake)
# NB-5：README 承诺「有需留意 / 异常项时结果会列出」→ 只有 warn 时不降级、退出码不变，但结论必须并进正文
check(r.kind == sa.OK and r.code == 0 and r.title == "一键安装完成", "install_all：健康只有 warn → 保持 OK / 0（不降级）")
check("■ 健康检查：1 项需留意" in r.body and "· user.version" in r.body and "点「一键安装」覆盖安装即可同步版本" in r.body,
      "install_all：健康只有 warn → 正文追加「■ 健康检查：N 项需留意：· id 消息（处置：提示）」")

def hc_boom(outq, heartbeat_wait_s=0): raise RuntimeError("hc boom")
real_hc = sa.health_check; sa.health_check = hc_boom
fake, _ = with_install_and_hc([0, 0], None, 0)
try:
    r, q = run_patched(sa.install_all, fake)
finally:
    sa.health_check = real_hc
check(r.kind == sa.OK and r.code == 0 and r.title == "一键安装完成", "install_all：健康检查自身出错 → 只记一行，安装结果与退出码不受影响")
check(any("健康检查" in str(x) and "出错" in str(x) for x in q.lines), "install_all：健康检查出错留了日志行")

# --- run_cli /healthcheck：退出码透传 + 日志快照（打桩，不写真实日志文件） ---
saved = {}
real_runner, real_append, real_save = sa.Runner, sa.append_log_file, sa.save_log_beside_exe
sa.append_log_file = lambda text: None
sa.save_log_beside_exe = lambda text, note="", **kw: saved.setdefault("called", True)   # **kw：run_cli 现在会传 target_dir=（交接模式）
fake, _ = with_install_and_hc([], HC_FAIL_LINES, 10)
sa.Runner = fake
try:
    with contextlib.redirect_stdout(io.StringIO()):
        rc = sa.run_cli("/healthcheck")
finally:
    sa.Runner, sa.append_log_file, sa.save_log_beside_exe = real_runner, real_append, real_save
check(rc == 10, "run_cli /healthcheck：退出码透传（脚本 10 → CLI 10）")
check(saved.get("called"), "run_cli /healthcheck：日志快照另存到 exe 旁被调用")



# ---------------- v1.4.5：检查更新 / 一键更新 ----------------
import sys as _sys
from unittest import mock

_su = sa.setup_update  # setup_actions 实际引用的更新模块对象（打桩要打在它身上）

def patch_update(fetch=None, run_update=None, replace=None):
    """给 setup_update 打桩（fetch_manifest / run_update / replace_old_exe），返回还原用的列表"""
    patches = []
    if fetch is not None:
        patches.append(mock.patch.object(_su, "fetch_manifest", fetch))
    if run_update is not None:
        patches.append(mock.patch.object(_su, "run_update", run_update))
    if replace is not None:
        patches.append(mock.patch.object(_su, "replace_old_exe", replace))
    return [p for p in [pm.__enter__() for pm in patches]], patches

def unpatch(entered, managers):
    for pm in managers:
        pm.__exit__(None, None, None)

MANIFEST_OK = {"version": "1.5.0", "url": "https://github.com/jasonisbrave/attachlink/releases/download/v1.5.0/AttachLink-1.5.0.zip",
               "sha256_zip": "a" * 64, "sha256_exe": "b" * 64, "size_zip": 100}

# check_update：有新版本 → PARTIAL / 20 + 固定格式的一行提示
entered, mgrs = patch_update(fetch=lambda mirror=None, log=print, manifest_url=None, current_version=None: (MANIFEST_OK, ""))
try:
    with mock.patch.object(sa, "read_version", lambda: "1.4.4"):
        r, q = run(lambda q: sa.check_update(q), [])
finally:
    unpatch(entered, mgrs)
check(r.kind == sa.PARTIAL and r.code == 20 and r.title == "发现新版本 v1.5.0",
      "check_update：有新版本 → PARTIAL / 20 / 发现新版本标题")
check(r.hint == "点「一键更新」即可自动下载并安装", "check_update：hint 指向一键更新")
check(any(str(x).strip() == "检查更新：最新 v1.5.0（当前 v1.4.4）" for x in q.lines),
      "check_update：日志含「检查更新：最新 vX（当前 vY）」固定行")
check(not any(str(x).startswith("发布说明：") for x in q.lines) and "发布说明" not in r.body,
      "check_update：清单没有 notes_url 时不输出发布说明行")
# 清单带 notes_url：body 与日志各带一行「发布说明：<地址>」（界面据此显示链接）
NOTES = "https://github.com/jasonisbrave/attachlink/releases/tag/v1.5.0"
entered, mgrs = patch_update(fetch=lambda mirror=None, log=print, manifest_url=None, current_version=None: (dict(MANIFEST_OK, notes_url=NOTES), ""))
try:
    with mock.patch.object(sa, "read_version", lambda: "1.4.4"):
        r, q = run(lambda q: sa.check_update(q), [])
finally:
    unpatch(entered, mgrs)
check(("\n发布说明：" + NOTES) in r.body and any(str(x).strip() == "发布说明：" + NOTES for x in q.lines),
      "check_update：清单带 notes_url → body 与日志各一行「发布说明：<地址>」")

# check_update：已是最新 → OK / 0（直连取得，文案不带缓存延迟说明）
entered, mgrs = patch_update(fetch=lambda mirror=None, log=print, manifest_url=None, current_version=None: (dict(MANIFEST_OK, version="1.4.4"), ""))
try:
    with mock.patch.object(sa, "read_version", lambda: "1.4.4"):
        r, q = run(lambda q: sa.check_update(q), [])
finally:
    unpatch(entered, mgrs)
check(r.kind == sa.OK and r.code == 0 and r.title == "已是最新版本" and "缓存延迟" not in r.body,
      "check_update：已是最新（直连）→ OK / 0，文案无缓存延迟说明")

# check_update：已是最新且经加速镜像取得（直连不可达）→ 文案注明可能有缓存延迟
entered, mgrs = patch_update(fetch=lambda mirror=None, log=print, manifest_url=None, current_version=None: (dict(MANIFEST_OK, version="1.4.4"), "https://gh-proxy.com/"))
try:
    with mock.patch.object(sa, "read_version", lambda: "1.4.4"):
        r, q = run(lambda q: sa.check_update(q), [])
finally:
    unpatch(entered, mgrs)
check(r.kind == sa.OK and "缓存延迟" in r.body and "Releases 页面手动下载" in r.body,
      "check_update：经镜像取得已是最新 → 文案注明可能有缓存延迟")

# check_update：清单拉取失败 → FAIL / 21，hint 含 Releases 手动下载入口
def fetch_boom(mirror=None, log=print, manifest_url=None, current_version=None):
    raise _su.UpdateError("network", "获取更新清单失败：…")
entered, mgrs = patch_update(fetch=fetch_boom)
try:
    with mock.patch.object(sa, "read_version", lambda: "1.4.4"):
        r, q = run(lambda q: sa.check_update(q), [])
finally:
    unpatch(entered, mgrs)
check(r.kind == sa.FAIL and r.code == 21 and "github.com/jasonisbrave/attachlink/releases" in r.hint,
      "check_update：失败 → FAIL / 21，hint 给出手动下载地址")

# apply_update：成功交接 → OK / hint=EXIT_AFTER（界面据此关闭自己）；【进度】行格式固定
def run_update_ok(**kw):
    assert kw.get("action") == "/install-all", "apply_update 默认动作应为 /install-all"
    kw["progress"](128, 256)
    kw["progress"](256, 256)
    kw["out"]("已启动新版本安装器（PID 99）")
    return "已启动新版本安装器（PID 99）"
entered, mgrs = patch_update(run_update=run_update_ok)
try:
    r, q = run(lambda q: sa.apply_update(q), [])
finally:
    unpatch(entered, mgrs)
check(r.kind == sa.OK and r.code == 0 and r.title == "已启动新版本安装器", "apply_update：交接成功 → OK")
check(r.hint == "EXIT_AFTER", "apply_update：成功时 hint = EXIT_AFTER（固定标记）")
check("本窗口即将关闭" in r.body and "管理员授权" in r.body, "apply_update：正文说明交接行为")
progress_lines = [str(x).rstrip("\n") for x in q.lines if str(x).startswith("【进度】")]
check(progress_lines == ["【进度】128\t256\t50%", "【进度】256\t256\t100%"],
      "apply_update：【进度】行格式 = 已下载\\t总字节\\t百分比%")

# apply_update：已是最新 → OK；失败 → FAIL / 22 / title 更新失败
entered, mgrs = patch_update(run_update=lambda **kw: "当前已是最新版本 v1.4.4")
try:
    r, _q = run(lambda q: sa.apply_update(q), [])
finally:
    unpatch(entered, mgrs)
check(r.kind == sa.OK and r.code == 0 and r.title == "已是最新版本", "apply_update：已是最新 → OK")
entered, mgrs = patch_update(run_update=lambda **kw: "失败：获取更新清单失败（network）建议：…")
try:
    r, _q = run(lambda q: sa.apply_update(q), [])
finally:
    unpatch(entered, mgrs)
check(r.kind == sa.FAIL and r.code == 22 and r.title == "更新失败" and r.body.startswith("失败："),
      "apply_update：失败 → FAIL / 22 / 更新失败")

# run_cli：识别 /check-update 与 /update，并把附加参数透传
class _Recorder:
    def __init__(self, ret): self.ret = ret; self.calls = []
    def __call__(self, *a, **kw): self.calls.append((a, kw)); return self.ret

saved = (sa.check_update, sa._apply_update, sa.save_log_beside_exe)
try:
    sa.save_log_beside_exe = lambda *a, **k: ""   # 不往源码目录写日志快照
    rec = _Recorder(sa.ActionResult(sa.OK, 0, "已是最新版本", "…"))
    sa.check_update = rec
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "/check-update"]):
        with contextlib.redirect_stdout(io.StringIO()):
            code = sa.run_cli("/check-update")
    check(code == 0 and rec.calls and rec.calls[0][0][1:] == (None, None),
          "run_cli：/check-update → 退出码 0，附加参数缺省为 None")
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "/check-update", "--mirror", "https://m/",
                                          "--manifest-url", "http://127.0.0.1:9/j"]):
        with contextlib.redirect_stdout(io.StringIO()):
            code = sa.run_cli("/check-update")
    check(rec.calls[-1][0][1:] == ("https://m/", "http://127.0.0.1:9/j"),
          "run_cli：--mirror / --manifest-url 透传给 check_update")

    rec2 = _Recorder(sa.ActionResult(sa.FAIL, 22, "更新失败", "失败：…"))
    sa._apply_update = rec2
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "/update", "--update-dry-run"]):
        with contextlib.redirect_stdout(io.StringIO()):
            code = sa.run_cli("/update")
    check(code == 22 and rec2.calls[0][1].get("dry_run") is True
          and rec2.calls[0][1].get("action") == "/install-all",
          "run_cli：/update → --update-dry-run 识别、默认动作 /install-all、退出码取 ActionResult.code")
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "/update"]):
        with contextlib.redirect_stdout(io.StringIO()):
            code = sa.run_cli("/update")
    check(rec2.calls[-1][1].get("dry_run") is False, "run_cli：/update 不带开关时 dry_run=False")
finally:
    sa.check_update, sa._apply_update, sa.save_log_beside_exe = saved

# run_cli：--updated-from 交接参数在任何动作前生效（这里以 /selfcheck 承载）
saved2 = (sa.setup_update.replace_old_exe,)
def _hand(ok, status, message, backup=None):
    return {"ok": ok, "status": status, "message": message, "backup": backup}
rec3 = _Recorder(_hand(True, "replaced", "旧安装器已替换为新版本"))
try:
    sa.setup_update.replace_old_exe = rec3
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "/selfcheck", "--updated-from", "C:\\old\\AttachLinkSetup.exe", "--wait-pid", "123"]):
        with contextlib.redirect_stdout(io.StringIO()):
            code = sa.run_cli("/selfcheck")
    check(rec3.calls and rec3.calls[0][0][0] == "C:\\old\\AttachLinkSetup.exe"
          and rec3.calls[0][1].get("wait_pid") == "123",
          "run_cli：--updated-from / --wait-pid 解析并交给 replace_old_exe")
    check(code == 3, "run_cli：交接后继续执行本行动作（/selfcheck 在无载荷环境返回 3）")
    # 没带 --updated-from 时不触发
    rec3.calls.clear()
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "/no-such-mode"]):
        with contextlib.redirect_stdout(io.StringIO()):
            sa.run_cli("/no-such-mode")
    check(not rec3.calls, "run_cli：无 --updated-from 不调用 replace_old_exe")
finally:
    sa.setup_update.replace_old_exe = saved2[0]

# ---------------- v1.4.5 补修：公开交接入口 / 快照目录 / 环境变量清单地址钩子 ----------------

# (补修3) 公开入口 handle_update_handoff：与旧名互为别名、行为一致
saved_h = (_su.replace_old_exe,)
try:
    rec_h = _Recorder(_hand(True, "replaced", "旧安装器已替换为新版本"))
    _su.replace_old_exe = rec_h
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "--updated-from", "C:\\old\\AttachLinkSetup.exe", "--wait-pid", "77"]):
        sa.handle_update_handoff()
    check(hasattr(sa, "_handle_updated_from") and sa._handle_updated_from is sa.handle_update_handoff,
          "handle_update_handoff：与旧名 _handle_updated_from 互为别名")
    check(rec_h.calls and rec_h.calls[0][0][0] == "C:\\old\\AttachLinkSetup.exe"
          and rec_h.calls[0][1].get("wait_pid") == "77",
          "handle_update_handoff：公开入口解析 --updated-from/--wait-pid 并调用 replace_old_exe")
    rec_h.calls.clear()
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "/selfcheck"]):
        sa.handle_update_handoff()
    check(not rec_h.calls, "handle_update_handoff：无 --updated-from 时不动作")
finally:
    _su.replace_old_exe = saved_h[0]

# (补修2) 交接模式日志快照写到旧安装器所在目录
with tempfile.TemporaryDirectory() as d:
    old_exe = os.path.join(d, "old", "AttachLinkSetup.exe")
    os.makedirs(os.path.dirname(old_exe))
    check(sa._log_beside_dir(["x", "--updated-from", old_exe, "--wait-pid", "1"]) == os.path.dirname(old_exe),
          "_log_beside_dir：--updated-from → 旧安装器所在目录")
    check(sa._log_beside_dir(["x", "/selfcheck"]) is None, "_log_beside_dir：无交接参数 → None")
    check(sa._log_beside_dir(["x", "--updated-from", os.path.join(d, "ghost", "a.exe")]) is None,
          "_log_beside_dir：旧目录不存在 → None（维持现状）")
    saved_s = (sa.check_update, sa.save_log_beside_exe)
    try:
        sa.check_update = _Recorder(sa.ActionResult(sa.OK, 0, "已是最新版本", "…"))
        with mock.patch.object(_sys, "argv", ["setup_gui.py", "/check-update", "--updated-from", old_exe]):
            with contextlib.redirect_stdout(io.StringIO()):
                code = sa.run_cli("/check-update")
        snap = os.path.join(os.path.dirname(old_exe), "AttachLinkSetup.log")
        check(code == 0 and os.path.isfile(snap)
              and "[silent /check-update" in open(snap, encoding="utf-8").read(),
              "run_cli：交接模式日志快照写到旧安装器所在目录")
    finally:
        sa.check_update, sa.save_log_beside_exe = saved_s

# ---------------- v1.5.0（oracle B 复核）：交接结果向上传播，界面 / CLI 如实呈现 ----------------

# B-2-1 handle_update_handoff 返回结构化结果：等待超时 / 覆盖失败 / 抛异常 三种失败都必须传播到返回值
def argv_with_handoff():
    return ["setup_gui.py", "/install-all", "--updated-from", "C:\\old dir\\AttachLinkSetup.exe", "--wait-pid", "123"]

saved_ho = _su.replace_old_exe
try:
    # 等待旧进程超时
    _su.replace_old_exe = lambda old, wait_pid=None: _hand(False, "failed", "等待旧安装器进程（PID 123）退出超时（15 秒），本次未替换")
    with mock.patch.object(_sys, "argv", argv_with_handoff()):
        hand = sa.handle_update_handoff()
    check(hand["attempted"] is True and hand["ok"] is False and "超时" in hand["message"]
          and hand["old_exe"] == "C:\\old dir\\AttachLinkSetup.exe" and hand["status"] == "failed",
          "handle_update_handoff：等待旧进程超时 → attempted/ok=False/message 含原因/status 传播（向上传播）")
    # 覆盖失败（杀毒锁定 / 目录不可写）
    _su.replace_old_exe = lambda old, wait_pid=None: _hand(False, "failed", "替换旧安装器失败（已回滚）：[Errno 13] Permission denied")
    with mock.patch.object(_sys, "argv", argv_with_handoff()):
        hand = sa.handle_update_handoff()
    check(hand["attempted"] is True and hand["ok"] is False and "已回滚" in hand["message"],
          "handle_update_handoff：覆盖失败 → ok=False（向上传播）")
    # 替换失败且回滚也失败：backup 路径必须透传
    _su.replace_old_exe = lambda old, wait_pid=None: _hand(False, "failed", "替换旧安装器失败：原位置目前没有安装器", "C:\\old dir\\AttachLinkSetup.exe.bak-abc123")
    with mock.patch.object(_sys, "argv", argv_with_handoff()):
        hand = sa.handle_update_handoff()
    check(hand["attempted"] is True and hand["ok"] is False
          and hand["backup"] == "C:\\old dir\\AttachLinkSetup.exe.bak-abc123" and hand["status"] == "failed",
          "handle_update_handoff：回滚也失败 → backup 路径透传（向上传播）")
    # 抛异常
    def boom_replace(old, wait_pid=None): raise RuntimeError("模拟交接异常")
    _su.replace_old_exe = boom_replace
    with mock.patch.object(_sys, "argv", argv_with_handoff()):
        hand = sa.handle_update_handoff()
    check(hand["attempted"] is True and hand["ok"] is False and "模拟交接异常" in hand["message"],
          "handle_update_handoff：replace_old_exe 抛异常 → ok=False，message 为异常说明（不向上抛）")
    # 成功
    _su.replace_old_exe = lambda old, wait_pid=None: _hand(True, "replaced", "旧安装器已替换为新版本")
    with mock.patch.object(_sys, "argv", argv_with_handoff()):
        hand = sa.handle_update_handoff()
    check(hand["attempted"] is True and hand["ok"] is True and "已替换" in hand["message"]
          and hand["status"] == "replaced" and hand["backup"] is None,
          "handle_update_handoff：替换成功 → ok=True，status/backup 传播")
    # 未尝试
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "/install-all"]):
        hand = sa.handle_update_handoff()
    check(hand["attempted"] is False and hand["ok"] is None and hand["message"] == ""
          and hand["status"] is None,
          "handle_update_handoff：无 --updated-from → attempted=False / ok=None")
finally:
    _su.replace_old_exe = saved_ho

# B-2-2 handoff_summary_line 纯函数：按 status 区分措辞
_hand_ok = {"attempted": True, "ok": True, "status": "replaced", "backup": None,
            "message": "旧安装器已替换为新版本", "old_exe": "C:\\old\\AttachLinkSetup.exe"}
_hand_fail = {"attempted": True, "ok": False, "status": "failed", "backup": None,
              "message": "等待旧安装器进程（PID 123）退出超时（15 秒），本次未替换",
              "old_exe": "C:\\old\\AttachLinkSetup.exe"}
check(sa.handoff_summary_line(None, "D:\\tmp\\x") == "" and sa.handoff_summary_line({"attempted": False}, "D:\\tmp\\x") == "",
      "handoff_summary_line：未尝试 → 空串")
check(sa.handoff_summary_line(_hand_ok, "D:\\tmp\\x") == "已由旧版本安装器交接启动，旧安装器文件已更新",
      "handoff_summary_line：replaced → 「旧安装器文件已更新」")
check(sa.handoff_summary_line(dict(_hand_ok, status="recreated"), "D:\\tmp\\x")
      == "已由旧版本安装器交接启动；原位置的安装器此前已不存在，已按新版本重建",
      "handoff_summary_line：recreated → 已按新版本重建")
check(sa.handoff_summary_line(dict(_hand_ok, status="same"), "D:\\tmp\\x")
      == "已由旧版本安装器交接启动（旧安装器即本安装器）",
      "handoff_summary_line：same → 旧安装器即本安装器")
check(sa.handoff_summary_line(dict(_hand_ok, ok=False, status="missing",
                                 message="旧安装器文件已不存在，且原位置不可写（Permission denied），未替换"),
                              "D:\\tmp\\guid").startswith(
      "已由旧版本安装器交接启动；旧安装器文件已不存在，且原位置不可写（Permission denied），未替换"
      "——本安装器位于 D:\\tmp\\guid，请把它放到固定位置"),
      "handoff_summary_line：missing → 如实说明未替换并给出本安装器位置")
_line = sa.handoff_summary_line(_hand_fail, "D:\\tmp\\guid")
check(_line.startswith("已由旧版本安装器交接启动；旧安装器文件未能替换（等待旧安装器进程（PID 123）退出超时")
      and "D:\\tmp\\guid" in _line and "复制到原来的安装器位置" in _line,
      "handoff_summary_line：failed → 如实说明未替换并给出本安装器位置与手动复制建议")
_line_bak = sa.handoff_summary_line(dict(_hand_fail, message="替换旧安装器失败：原位置目前没有安装器",
                                         backup="C:\\old\\AttachLinkSetup.exe.bak-abc123"), "D:\\tmp\\guid")
check("旧文件的备份保留在：C:\\old\\AttachLinkSetup.exe.bak-abc123" in _line_bak
      and "原位置目前没有安装器，请把该备份改回原名" in _line_bak,
      "handoff_summary_line：failed 且 backup 非空 → 追加备份路径与改回提示")

# B-2-3 run_cli：交接失败不改退出码，但 [结果] 行与最终打印行必须追加提示
saved_cli = (sa.check_update, sa.append_log_file, sa.save_log_beside_exe, _su.replace_old_exe)
logged = []
try:
    sa.check_update = lambda q, mirror=None, manifest_url=None: sa.ActionResult(sa.OK, 0, "已是最新版本", "…")
    sa.append_log_file = lambda text: logged.append(text)
    sa.save_log_beside_exe = lambda *a, **k: ""
    _su.replace_old_exe = lambda old, wait_pid=None: _hand(False, "failed", "替换旧安装器失败（已回滚）：被占用")
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "/check-update", "--updated-from", "C:\\old\\AttachLinkSetup.exe"]):
        with contextlib.redirect_stdout(io.StringIO()) as out_buf:
            code = sa.run_cli("/check-update")
    check(code == 0, "run_cli：交接失败 → 退出码不变（安装动作本身可能成功）")
    check(any("；注意：旧安装器文件未替换：替换旧安装器失败（已回滚）：被占用" in t for t in logged),
          "run_cli：交接失败 → 日志 [结果] 行追加「；注意：旧安装器文件未替换：<原因>」")
    check("；注意：旧安装器文件未替换：替换旧安装器失败（已回滚）：被占用" in out_buf.getvalue(),
          "run_cli：交接失败 → 最终打印行同样追加提示")
    # 交接成功不追加提示
    logged.clear()
    _su.replace_old_exe = lambda old, wait_pid=None: _hand(True, "replaced", "旧安装器已替换为新版本")
    with mock.patch.object(_sys, "argv", ["setup_gui.py", "/check-update", "--updated-from", "C:\\old\\AttachLinkSetup.exe"]):
        with contextlib.redirect_stdout(io.StringIO()):
            code = sa.run_cli("/check-update")
    check(code == 0 and not any("未替换" in t for t in logged), "run_cli：交接成功 → 无提示追加")
finally:
    sa.check_update, sa.append_log_file, sa.save_log_beside_exe, _su.replace_old_exe = saved_cli

# (补修4) 环境变量 ATTACHLINK_UPDATE_MANIFEST_URL：仅回环地址生效
_env = "ATTACHLINK_UPDATE_MANIFEST_URL"
saved_f, saved_ru = _su.fetch_manifest, _su.run_update
try:
    _su.fetch_manifest = _Recorder((MANIFEST_OK, ""))
    with mock.patch.dict(os.environ, {_env: "http://127.0.0.1:9/j"}):
        with mock.patch.object(sa, "read_version", lambda: "1.4.4"):
            r, q = run(lambda q: sa.check_update(q), [])
    check(_su.fetch_manifest.calls[0][1].get("manifest_url") == "http://127.0.0.1:9/j",
          "check_update：环境变量（回环）→ 作为默认清单地址传给 fetch_manifest")
    check(_su.fetch_manifest.calls[0][1].get("current_version") == "1.4.4",
          "check_update：当前版本传给 fetch_manifest（供直连复核判断）")
    with mock.patch.dict(os.environ, {_env: "https://example.com/latest.json"}):
        with mock.patch.object(sa, "read_version", lambda: "1.4.4"):
            r, q = run(lambda q: sa.check_update(q), [])
    check(_su.fetch_manifest.calls[-1][1].get("manifest_url") is None
          and any("已忽略" in str(x) for x in q.lines),
          "check_update：环境变量（非回环）→ 忽略并写一行日志说明")
    os.environ.pop(_env, None)
    with mock.patch.object(sa, "read_version", lambda: "1.4.4"):
        r, q = run(lambda q: sa.check_update(q), [])
    check(_su.fetch_manifest.calls[-1][1].get("manifest_url") is None,
          "check_update：未设置环境变量 → None")

    rec_ru = _Recorder("已启动新版本安装器（PID 5）")
    _su.run_update = rec_ru
    with mock.patch.dict(os.environ, {_env: "http://localhost:9/j"}):
        r, q = run(lambda q: sa.apply_update(q), [])
    check(rec_ru.calls[0][1].get("manifest_url") == "http://localhost:9/j",
          "apply_update：环境变量（回环 localhost）→ 传给 run_update")
    with mock.patch.dict(os.environ, {_env: "ftp://example.com/x"}):
        r, q = run(lambda q: sa.apply_update(q), [])
    check(rec_ru.calls[-1][1].get("manifest_url") is None and any("已忽略" in str(x) for x in q.lines),
          "apply_update：环境变量（非 http/https）→ 忽略并写日志说明")
finally:
    _su.fetch_manifest = saved_f
    _su.run_update = saved_ru
    os.environ.pop(_env, None)

# ---------------- v1.5.0（council 修复批 A）：卸载 / 部署退出码映射与残留清单 ----------------

RESIDUE_LINES = [
    "⚠ 卸载未完全成功，以下项目未能清除：",
    "  - 加载项副本：C:\\Users\\u\\AppData\\Roaming\\Microsoft\\AddIns\\AttachLink.xlam（文件未能删除）",
    "  - 受信任位置登记：HKCU:\\…\\15.0\\…\\AttachLink（注册表键未能删除）",
    "请关闭所有 Excel（含鼎信诺、WPS 表格）后重新运行卸载；仍失败则按上方清单手动处理。",
]

check(sa._extract_residue_lines(RESIDUE_LINES) == [
    "加载项副本：C:\\Users\\u\\AppData\\Roaming\\Microsoft\\AddIns\\AttachLink.xlam（文件未能删除）",
    "受信任位置登记：HKCU:\\…\\15.0\\…\\AttachLink（注册表键未能删除）",
], "_extract_residue_lines：从子进程输出截取「- 」残留行，遇普通行即停止")

class _ScriptRunner:
    """假 Runner：run / run_elevated 依次返回 codes；run 先把预设输出行进 outq（供残留清单截取）"""
    def __init__(self, codes, lines=None):
        self.seq = {"run": iter([c for k, c in codes if k == "S"]), "run_elevated": iter([c for k, c in codes if k == "E"])}
        self.lines = lines or []
    def make(self):
        seq, lines = self.seq, self.lines
        class R:
            def __init__(self, outq): self.outq = outq
            def run(self, args, cwd):
                for ln in lines: self.outq.put(ln)
                return next(seq["run"])
            def run_elevated(self, target): return next(seq["run_elevated"])
        return R

# 打桩收尾清理，避免测试真去删 %TEMP% 里的看门狗日志 / 更新临时目录
_saved_cleanup = sa._uninstall_cleanup_tmp
sa._uninstall_cleanup_tmp = lambda: None
try:
    fake = _ScriptRunner([("S", 2)]).make()
    r, q = run_patched(sa.uninstall, fake)
    check(r.kind == sa.FAIL and r.code == 2 and "卸载未执行" in r.title and "加载项文件被占用" in r.title,
          "卸载 exit 2 → FAIL / 2 /「卸载未执行：Excel 正在运行或加载项文件被占用」")

    fake = _ScriptRunner([("S", 10)], RESIDUE_LINES).make()
    r, q = run_patched(sa.uninstall, fake)
    check(r.kind == sa.FAIL and r.code == 10 and "有残留" in r.title, "卸载 exit 10 → FAIL / 10 /「有残留」")
    check("加载项副本" in r.body and "受信任位置登记" in r.body, "卸载 exit 10 → 正文截入子进程输出的残留清单")

    fake = _ScriptRunner([("S", 9)]).make()
    r, q = run_patched(sa.uninstall, fake)
    check(r.kind == sa.FAIL and r.code == 9, "卸载 exit 9 → FAIL / 9（既有契约不变）")

    fake = _ScriptRunner([("S", 0)]).make()
    r, q = run_patched(sa.uninstall, fake)
    check(r.kind == sa.OK and r.code == 0 and "核验无残留" in r.body, "卸载 exit 0 → OK，正文如实带「按事实核验无残留」")

    fake = _ScriptRunner([("E", 12)]).make()
    r, q = run_patched(sa.deploy_elevated, fake)
    check(r.kind == sa.FAIL and r.code == 12 and "其他管理员账户" in r.title,
          "单独部署 exit 12 → FAIL / 12 /「使用了其他管理员账户授权」")

    fake = _ScriptRunner([("E", 0)]).make()
    r, q = run_patched(sa.deploy_elevated, fake)
    check(r.kind == sa.OK and "已部署并请求启动" in r.body and "15 秒后的健康检查" in r.body,
          "部署成功文案：「已部署并请求启动（15 秒后的健康检查会核对其心跳）」，不再先说「已启动」")

    fake = _ScriptRunner([("E", 8)]).make()
    r, q = run_patched(sa.uninstall_elevated_only, fake)
    check(r.kind == sa.FAIL and r.code == 8 and "未完全成功" in r.title, "机器级卸载 exit 8 → FAIL / 8 / 未完全成功")
    fake = _ScriptRunner([("E", 0)]).make()
    r, q = run_patched(sa.uninstall_elevated_only, fake)
    check(r.kind == sa.OK and "核验无残留" in r.body, "机器级卸载 exit 0 → OK，正文如实带核验表述")
finally:
    sa._uninstall_cleanup_tmp = _saved_cleanup

fake, _hc = with_install_and_hc([0, 12], HC_OK_LINES, 0)
r, q = run_patched(sa.install_all, fake)
check(r.kind == sa.PARTIAL and r.code == 12 and "其他管理员账户" in r.title,
      "install_all 第二步 exit 12 → PARTIAL / 12 /「使用了其他管理员账户授权」文案")



# ---------------- v1.6.0（R3）：图形界面日志快照写到旧安装器所在目录 ----------------
# 缺陷：界面三处收尾 save_log_beside_exe(...) 没传 target_dir，交接模式下快照落进 %TEMP% 更新
# 临时目录（用户看不到），而不是旧安装器所在目录。修复后统一走模块级 save_gui_log_snapshot。
import ast as _ast

_HERE_PACK = os.path.dirname(os.path.abspath(__file__))
if _HERE_PACK not in _sys.path:
    _sys.path.insert(0, _HERE_PACK)   # setup_gui 的 import setup_actions / setup_update 需要
_sys.modules.setdefault("setup_actions", sa)   # 复用同一模块对象（打桩 sa.exe_dir 才生效）
_spec_gui = importlib.util.spec_from_file_location("setup_gui", os.path.join(_HERE_PACK, "setup_gui.py"))
sg = importlib.util.module_from_spec(_spec_gui)
_sys.modules.setdefault("setup_gui", sg)
_spec_gui.loader.exec_module(sg)   # 模块级只有常量与函数定义；Tk 窗口在 App.__init__，import 不开窗

with tempfile.TemporaryDirectory() as d:
    old_exe = os.path.join(d, "old", "AttachLinkSetup.exe")
    os.makedirs(os.path.dirname(old_exe))
    open(old_exe, "w").close()
    exe_dir_tmp = os.path.join(d, "exedir")
    os.makedirs(exe_dir_tmp)
    real_exe_dir = sa.exe_dir
    sa.exe_dir = lambda: exe_dir_tmp   # 源码态 exe_dir() 指向 packaging/ 源码目录，打桩到临时目录
    try:
        # (a) 交接模式：--updated-from <旧exe> → 快照落在旧安装器所在目录且内容一致
        with mock.patch.object(_sys, "argv", ["setup_gui.py", "--updated-from", old_exe, "--wait-pid", "7"]):
            got = sg.save_gui_log_snapshot("gui-log-text", "note-1")
        snap_old = os.path.join(os.path.dirname(old_exe), "AttachLinkSetup.log")
        check(got == snap_old and os.path.isfile(snap_old)
              and open(snap_old, encoding="utf-8").read() == "gui-log-text\nnote-1\n",
              "save_gui_log_snapshot：交接模式快照写到旧安装器所在目录且内容一致")
        check(not os.path.exists(os.path.join(exe_dir_tmp, "AttachLinkSetup.log")),
              "save_gui_log_snapshot：交接模式不写 exe 所在目录")
        os.remove(snap_old)
        # (b) 无 --updated-from → 回落 exe 所在目录
        with mock.patch.object(_sys, "argv", ["setup_gui.py"]):
            got = sg.save_gui_log_snapshot("plain-log")
        snap_exe = os.path.join(exe_dir_tmp, "AttachLinkSetup.log")
        check(got == snap_exe and os.path.isfile(snap_exe)
              and open(snap_exe, encoding="utf-8").read() == "plain-log",
              "save_gui_log_snapshot：无交接参数 → 快照写 exe 所在目录")
        os.remove(snap_exe)
        # (c) --updated-from 指向不存在的目录 → 回落 exe 所在目录
        with mock.patch.object(_sys, "argv", ["setup_gui.py", "--updated-from", os.path.join(d, "ghost", "a.exe")]):
            got = sg.save_gui_log_snapshot("fallback-log")
        check(got == snap_exe and os.path.isfile(snap_exe)
              and open(snap_exe, encoding="utf-8").read() == "fallback-log",
              "save_gui_log_snapshot：旧目录不存在 → 回落 exe 所在目录")
    finally:
        sa.exe_dir = real_exe_dir

# (d) 静态守卫：setup_gui.py 里每一处 save_log_beside_exe( 调用都必须带 target_dir 且在辅助函数内
_gui_src = open(os.path.join(_HERE_PACK, "setup_gui.py"), encoding="utf-8").read()
check('if __name__ == "__main__":' in _gui_src,
      "setup_gui：有 __main__ 守卫（import 不创建 Tk 窗口）")

class _CallGuard(_ast.NodeVisitor):
    """记录每一处 save_log_beside_exe( 调用的（外层函数名, 是否带 target_dir）"""
    def __init__(self):
        self.stack = []
        self.calls = []
    def visit_FunctionDef(self, node):
        self.stack.append(node.name)
        self.generic_visit(node)
        self.stack.pop()
    def visit_Call(self, node):
        if isinstance(node.func, _ast.Name) and node.func.id == "save_log_beside_exe":
            self.calls.append((self.stack[-1] if self.stack else "<module>",
                               any(k.arg == "target_dir" for k in node.keywords)))
        self.generic_visit(node)

_guard = _CallGuard()
_guard.visit(_ast.parse(_gui_src))
check(len(_guard.calls) >= 1, "setup_gui：save_log_beside_exe 仍被调用（经 save_gui_log_snapshot）")
check(all(fn == "save_gui_log_snapshot" and has_target for fn, has_target in _guard.calls),
      f"setup_gui：每一处 save_log_beside_exe( 都在 save_gui_log_snapshot 内且带 target_dir（实际 { _guard.calls }）")
check(hasattr(sg, "save_gui_log_snapshot"), "setup_gui：模块级辅助函数 save_gui_log_snapshot 存在")



print("RESULT:", "PASS" if ok else "FAIL")

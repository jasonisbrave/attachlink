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

check(set(sa.CLI_ACTIONS) == {"/install-all", "/install", "/deploy-elevated", "/uninstall", "/uninstall-elevated"}, "静默参数表完整")

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
            for n in ["dist\\AttachLink.xlam", "scripts\\setup.ps1", "scripts\\watchdog.ps1", "scripts\\watchdog_launcher.vbs",
                      "build\\install.ps1", "build\\install-elevated.ps1", "run-elevated.ps1", "elevated-capture.ps1"]:
                os.makedirs(os.path.dirname(os.path.join(d, "payload", n)), exist_ok=True)
                open(os.path.join(d, "payload", n), "w").close()
            with contextlib.redirect_stdout(io.StringIO()):
                check(sa.selfcheck() == 0, "selfcheck：载荷齐全且版本可读 → 0")
            open(os.path.join(d, "payload", "VERSION.txt"), "w").close()   # 版本文件变空
            with contextlib.redirect_stdout(io.StringIO()):
                check(sa.selfcheck() == 3, "selfcheck：VERSION.txt 为空 → 3（不再静默显示未知版本）")
    finally:
        sa.payload_dir = real_payload_dir

# run_cli 的纯逻辑分支（不触发安装）：未知参数 → 1
with contextlib.redirect_stdout(io.StringIO()):
    check(sa.run_cli("/no-such-mode") == 1, "未知静默参数 → 1")

print("RESULT:", "PASS" if ok else "FAIL")

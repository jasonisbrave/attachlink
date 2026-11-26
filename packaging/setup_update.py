# -*- coding: utf-8 -*-
"""
AttachLink 安装器 —— 一键更新（v1.5.0 起）
更新链路：取发布清单 latest.json（GitHub Releases「最新版本」稳定地址，可经加速镜像前缀回退）
→ 与当前版本比较 → 下载发布 zip（状态 / 类型 / 魔数 / 大小 / SHA-256 全部校验，逐镜像回退）
→ 解压出顶层唯一的 AttachLinkSetup.exe（防 zip-slip，SHA-256 复验）
→ dry_run 到此为止；否则把新安装器拉起并携带 --updated-from / --wait-pid 交接参数，
由新安装器替换旧 exe 后自行执行安装动作。

信任模型：加速镜像只是通道，不改变内容；完整性以清单里的 SHA-256 为准（镜像与直连取得
的字节都必须与清单一致才算通过）。清单本身经 https + 固定主机 + 固定路径前缀白名单约束。
`--manifest-url` 指向回环地址时进入本地测试模式：允许 http，且放宽清单内下载地址的白名单
（仅此场景，正常路径的门槛不变）。只用 Python 标准库，不引入第三方包。
"""
import base64
import hashlib
import hmac
import http.client
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile

REPO = "jasonisbrave/attachlink"
MANIFEST_URL = f"https://github.com/{REPO}/releases/latest/download/latest.json"
MIRROR_PREFIXES = ["https://gh-proxy.com/", "https://ghfast.top/", "https://ghproxy.net/"]
ALLOWED_DOWNLOAD_HOST = "github.com"
ALLOWED_DOWNLOAD_PATH_PREFIX = f"/{REPO}/releases/download/"
# GitHub 会把 Release 资产（清单与安装包）302 到下列主机。只有当请求目标确实是本仓库的 Release 资产地址、
# 或是内置加速镜像时，才允许最终地址落在这些主机上（v1.5.1：此前不在白名单，GitHub 直连取清单一律被拦下）
GITHUB_ASSET_REDIRECT_HOSTS = {
    "release-assets.githubusercontent.com",
    "objects.githubusercontent.com",
    "github-releases.githubusercontent.com",
}
HTTP_TIMEOUT = 20          # 秒
MANIFEST_TIMEOUT = 8       # 秒
MAX_PACKAGE_BYTES = 200 * 1024 * 1024
MAX_MANIFEST_BYTES = 64 * 1024          # 清单大小上限：清单只是几行 JSON，超过即判「清单异常大」
MANIFEST_TOTAL_TIMEOUT = 30             # 秒：整轮清单获取（所有前缀合计）总时限，超时不再尝试剩余前缀

# 发布密钥（RSA 2048，RSASSA-PKCS#1 v1.5 + SHA-256）：内嵌公钥只用于验签 latest.json；
# 私钥（%USERPROFILE%\.attachlink\release-key.xml）不入库、不进安装包，由 build\new-release-key.ps1 生成。
# 更换密钥时同步替换 RELEASE_KEY_N（new-release-key.ps1 会打印新公钥与指纹），并用 --verify-manifest 复核。
# 换钥记录：2026-11-26（v1.6.0 开发批次，该批次并入 2.0 分支、不单独发版）——原私钥文件丢失后重建，公钥指纹
# e19ce2930c924fb6。v1.5.0 / v1.5.1 安装器内嵌的是旧公钥，验不过本密钥签发的清单（表现为「清单签名校验失败」、
# 退出码 21，不会误装），需手动下载一次带新公钥的版本（2.0 系）安装器。
RELEASE_KEY_N = (
    "f1d9b7e3849b40e89a0ee77691584069b6d0833224f92011a41e9b841a1bdb05"
    "873550ac9907baf916c0940410be8168e6639258398f5ae1d9d82f986d8b2abb"
    "30f7329835e956bbaba0621266a221798c357da21eaecacf754585c98a379e21"
    "e9e15315a53ff3800621a47734a8819ae925f02a871cc1831b8e16f192ae4d8b"
    "86e77498389e49d15d97dd59230a0baffa9114ca29452003840b80f6f15a8ab2"
    "dd0dfd46a221d2e916811587fbcf1d4daaaff04960c2bfb0f8e91996c980673e"
    "3df9ebd38e3181c2b74691421178d1d258db8756a8a516180c8bc686a6d06b81"
    "2f0f16f8715539d0c5b9d539af34e326b1d4a329a95e3957439de5d225b11dc9"
)
RELEASE_KEY_E = 65537
SIG_ALG = "RS256-SHA256-PKCS1v15:v1"
_SIG_DIGESTINFO_PREFIX = bytes.fromhex("3031300d060960864801650304020105000420")   # SHA-256 的 DER DigestInfo 前缀（19 字节）

_TMP_DIR_NAME = "AttachLinkSetup_update"
_INSTALLER_NAME = "AttachLinkSetup.exe"
_HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
_LOOPBACK_HOSTS = {"127.0.0.1", "localhost"}

# 各失败类别给用户看的建议（run_update 把 UpdateError 转成「失败：…」文本时拼接）
_SUGGESTIONS = {
    "network": "建议：检查网络后重试，或到 https://github.com/jasonisbrave/attachlink/releases 手动下载。",
    "manifest": "建议：稍后重试；仍失败请到 https://github.com/jasonisbrave/attachlink/releases 手动下载。",
    "verify": "建议：重新点一次「一键更新」；仍失败请到 Releases 页手动下载。",
    "extract": "建议：重新更新一次；仍失败请到 Releases 页手动下载。",
    "handoff": "建议：手动运行下载目录里的新安装器（路径见日志）。",
}

# 上一次成功使用的加速前缀（fetch_manifest / download_package 成功后记录，下次优先）
_last_good_prefix = ""


# ---------------- 版本 ----------------

def parse_version(s):
    """把「1.4.5」解析成整数元组 (1, 4, 5)；格式非法返回 None。"""
    if not isinstance(s, str):
        return None
    s = s.strip()
    if not s:
        return None
    parts = s.split(".")
    if any(not p.isdigit() for p in parts):
        return None
    return tuple(int(p) for p in parts)


def is_newer(a, b):
    """a 版本严格大于 b 时返回 True；任一版本无法解析返回 False（宁可少报不误报）。"""
    va, vb = parse_version(a), parse_version(b)
    if va is None or vb is None:
        return False
    return va > vb


# ---------------- 错误 ----------------

class UpdateError(Exception):
    """更新过程的可控失败：code ∈ network / manifest / verify / extract / handoff，message 为中文说明"""

    def __init__(self, code, message):
        super().__init__(message)
        self.code = code
        self.message = message


class _VerifyFailure(Exception):
    """下载 / 解压途中发现的完整性问题（内部用，最终转成 UpdateError("verify")）"""

    def __init__(self, reason):
        super().__init__(reason)
        self.reason = reason


# ---------------- HTTP ----------------

def _current_version() -> str:
    """当前安装器版本（供 User-Agent）；延迟导入避免与 setup_actions 循环依赖"""
    try:
        from setup_actions import read_version  # 延迟导入
        return read_version()
    except Exception:  # noqa: BLE001  单独加载本模块（单测）时取不到，退化为通用 UA
        return "未知版本"


def _user_agent() -> str:
    # 请求头只允许 ASCII：版本号读不到 / 非法时用 unknown，否则 urllib 会因 latin-1 编码失败拒绝整个请求
    v = _current_version()
    return f"AttachLinkSetup/{v}" if parse_version(v) is not None else "AttachLinkSetup/unknown"


def _check_url_scheme(url: str, allow_loopback_http: bool) -> None:
    """只允许 https；显式放行时也仅限回环地址的 http（本地测试用）"""
    parts = urllib.parse.urlsplit(url)
    if parts.scheme == "https":
        return
    if (parts.scheme == "http" and allow_loopback_http
            and (parts.hostname or "").lower() in _LOOPBACK_HOSTS):
        return
    raise UpdateError("network", f"不支持的地址（仅允许 https{'，回环地址可 http' if allow_loopback_http else ''}）：{url}")


def _is_loopback_url(url: str) -> bool:
    """地址主机是否为回环地址（决定本地测试是否放宽下载地址门槛）"""
    try:
        host = (urllib.parse.urlsplit(url).hostname or "").lower()
    except ValueError:
        return False
    return host in _LOOPBACK_HOSTS


def _open_request(url: str, timeout: float):
    """带 User-Agent 打开 GET 响应；HTTP 错误与网络异常统一转 UpdateError("network")"""
    request = urllib.request.Request(url, headers={"User-Agent": _user_agent()})
    try:
        return urllib.request.urlopen(request, timeout=timeout)
    except urllib.error.HTTPError as e:
        raise UpdateError("network", f"HTTP {e.code}：{url}") from e
    except Exception as e:  # noqa: BLE001  超时 / DNS / 代理 / SSL 等一律按网络失败
        raise UpdateError("network", f"请求失败：{url}（{e}）") from e


def _resp_fields(resp) -> tuple:
    """从响应对象取 (状态码, Content-Type)；兼容打桩对象"""
    status = int(getattr(resp, "status", 0) or resp.getcode() or 0)
    content_type = (resp.headers.get("Content-Type") or "").strip()
    return status, content_type


def _is_github_release_asset_url(url: str) -> bool:
    """是否本仓库在 GitHub 上的 Release 资产地址：安装包（releases/download/…）或清单（releases/latest/download/…）。"""
    try:
        parts = urllib.parse.urlsplit(url)
    except ValueError:
        return False
    if (parts.scheme or "").lower() != "https" or (parts.hostname or "").lower() != ALLOWED_DOWNLOAD_HOST:
        return False
    return (parts.path.startswith(ALLOWED_DOWNLOAD_PATH_PREFIX)
            or parts.path.startswith(f"/{REPO}/releases/latest/download/"))


def _mirror_targets_release_asset(requested_url: str) -> bool:
    """请求地址是否为「内置加速镜像前缀 + 本仓库的 Release 资产地址」。
    只看请求主机是不是镜像还不够：镜像可以代理任意地址（如 <镜像>https://evil.example/x），
    那种请求不应获得 GitHub 资产主机的放行。"""
    for prefix in MIRROR_PREFIXES:
        if requested_url.startswith(prefix):
            return _is_github_release_asset_url(requested_url[len(prefix):])
    return False


def _final_url_allowed(final_url: str, requested_url: str, allow_loopback_http: bool) -> bool:
    """最终地址（urllib 自动跟随重定向之后）白名单：https 且主机 ∈ {github.com, 内置镜像主机, 请求目标主机}；
    请求目标是本仓库的 Release 资产地址、或「内置镜像 + 本仓库 Release 资产地址」时，额外放行 GitHub 的资产重定向主机；
    本地测试模式（allow_loopback_http）额外放行回环主机的 http/https。除此之外一律不可信。
    只检查最终地址，不逐跳检查中间重定向（内容另有签名与哈希校验兜底）。"""
    try:
        parts = urllib.parse.urlsplit(final_url)
    except ValueError:
        return False
    scheme = (parts.scheme or "").lower()
    host = (parts.hostname or "").lower()
    if scheme == "http" and allow_loopback_http and host in _LOOPBACK_HOSTS:
        return True
    if scheme != "https":
        return False
    hosts = {ALLOWED_DOWNLOAD_HOST}
    mirror_hosts = {(urllib.parse.urlsplit(p).hostname or "").lower() for p in MIRROR_PREFIXES}
    hosts |= mirror_hosts
    try:
        requested_host = (urllib.parse.urlsplit(requested_url).hostname or "").lower()
    except ValueError:
        requested_host = ""
    hosts.add(requested_host)                    # 调用方显式指定的前缀主机
    # GitHub 直连与内置加速镜像都可能把 Release 资产 302 到 GitHub 的资产主机：仅当请求目标确实是
    # 本仓库的 Release 资产地址、或是「内置镜像 + 本仓库资产地址」时才放行这一跳（别的来源跳到资产主机仍然拒绝）
    if _is_github_release_asset_url(requested_url) or _mirror_targets_release_asset(requested_url):
        hosts |= GITHUB_ASSET_REDIRECT_HOSTS
    if allow_loopback_http:
        hosts |= _LOOPBACK_HOSTS
    return host in hosts


def _check_final_url(resp, requested_url: str, allow_loopback_http: bool) -> None:
    """拿到响应后复查最终地址：被 302 降级到 http 或跳到白名单外主机 → 立即拒绝（不读内容）。
    打桩对象没有 geturl() 时无从复查，维持现状。"""
    geturl = getattr(resp, "geturl", None)
    if not callable(geturl):
        return
    final = geturl() or requested_url
    if not _final_url_allowed(final, requested_url, allow_loopback_http):
        raise UpdateError("manifest", f"最终地址未通过校验（疑似被重定向到不可信地址）：{final}")


def _http_get(url, timeout, on_chunk=None, allow_loopback_http=False, max_bytes=None):
    """GET 一个地址，返回 (状态码, Content-Type, 响应体字节)；on_chunk(n) 每收到一块回调一次。
    max_bytes：响应体大小上限（清单用；清单只是几行 JSON，超过即抛 manifest 错误「清单异常大」）。"""
    _check_url_scheme(url, allow_loopback_http)
    with _open_request(url, timeout) as resp:
        _check_final_url(resp, url, allow_loopback_http)
        status, content_type = _resp_fields(resp)
        chunks = []
        total = 0
        try:
            while True:
                block = resp.read(65536)
                if not block:
                    break
                chunks.append(block)
                total += len(block)
                if max_bytes is not None and total > max_bytes:
                    raise UpdateError("manifest", f"清单异常大（超过 {max_bytes} 字节），已拒绝读取")
                if on_chunk:
                    on_chunk(len(block))
        except UpdateError:
            raise
        except (OSError, http.client.HTTPException) as e:
            # 响应头已拿到（200）但读取响应体中途断流（超时 / 连接重置 / IncompleteRead 等）：
            # 按网络失败处理，让 fetch_manifest 换下一个来源，而不是整轮更新直接异常结束
            raise UpdateError("network", f"读取响应中断：{url}（{e}）") from e
        return status, content_type, b"".join(chunks)


def _http_download(url, timeout, write, on_response=None, on_chunk=None, allow_loopback_http=False):
    """流式 GET：响应头先经 on_response(状态, 类型, 总长) 校验（可抛 _VerifyFailure 提前终止），
    然后逐块 write(block) 落盘并 on_chunk(已收字节, 总长或 -1) 上报；返回 (状态码, Content-Type)"""
    _check_url_scheme(url, allow_loopback_http)
    with _open_request(url, timeout) as resp:
        status, content_type = _resp_fields(resp)
        length = (resp.headers.get("Content-Length") or "").strip()
        total = int(length) if length.isdigit() else -1
        if on_response:
            on_response(status, content_type, total)
        done = 0
        if on_chunk:
            on_chunk(done, total)
        try:
            while True:
                block = resp.read(256 * 1024)
                if not block:
                    break
                write(block)
                done += len(block)
                if on_chunk:
                    on_chunk(done, total)
        except _VerifyFailure:
            raise
        except UpdateError:
            raise
        except (OSError, http.client.HTTPException) as e:
            # 边下边读中途断流（超时 / 连接重置 / IncompleteRead 等）：转网络失败，
            # 让 download_package 删 .part 换下一个来源
            raise UpdateError("network", f"下载中断：{url}（{e}）") from e
        return status, content_type


# ---------------- 清单 ----------------

def validate_manifest(m, relax_url_gate=False):
    """校验清单字段（version / url / 两个 SHA-256 / size_zip / notes_url）；非法抛 UpdateError("manifest")。
    relax_url_gate 仅供本地回环测试清单使用：url 放宽为「https 或回环 http」，其余校验不变。"""
    if not isinstance(m, dict):
        raise UpdateError("manifest", "清单必须是 JSON 对象")
    if parse_version(m.get("version")) is None:
        raise UpdateError("manifest", f"清单 version 非法：{m.get('version')!r}")
    url = m.get("url")
    if not isinstance(url, str) or not url:
        raise UpdateError("manifest", "清单缺少 url")
    parts = urllib.parse.urlsplit(url)
    if not parts.path.endswith(".zip"):
        raise UpdateError("manifest", f"清单 url 必须以 .zip 结尾：{url}")
    if relax_url_gate:
        loopback_http = parts.scheme == "http" and (parts.hostname or "").lower() in _LOOPBACK_HOSTS
        if parts.scheme != "https" and not loopback_http:
            raise UpdateError("manifest", f"清单 url 协议非法（仅 https 或本地回环 http）：{url}")
    else:
        if parts.scheme != "https":
            raise UpdateError("manifest", f"清单 url 必须是 https：{url}")
        if (parts.hostname or "").lower() != ALLOWED_DOWNLOAD_HOST:
            raise UpdateError("manifest", f"清单 url 主机必须是 {ALLOWED_DOWNLOAD_HOST}：{url}")
        if not parts.path.startswith(ALLOWED_DOWNLOAD_PATH_PREFIX):
            raise UpdateError("manifest", f"清单 url 必须指向发布资产（{ALLOWED_DOWNLOAD_PATH_PREFIX}…）：{url}")
    for key in ("sha256_zip", "sha256_exe"):
        v = m.get(key)
        if not isinstance(v, str) or not _HEX64.match(v):
            raise UpdateError("manifest", f"清单 {key} 必须是 64 位十六进制")
    size = m.get("size_zip")
    if size is not None:
        if isinstance(size, bool) or not isinstance(size, int) or size < 0:
            raise UpdateError("manifest", f"清单 size_zip 必须是非负整数：{size!r}")
    notes = m.get("notes_url")
    if notes is not None:
        if not isinstance(notes, str) or urllib.parse.urlsplit(notes).scheme != "https":
            raise UpdateError("manifest", f"清单 notes_url 必须是 https 地址：{notes!r}")


# ---------------- 清单签名 ----------------

def manifest_signing_text(m: dict) -> str:
    """按签名规范生成被签名的正文（canonical text）：
    固定行序、UTF-8 无 BOM、\\n 换行、行尾也带 \\n；size_zip 缺失整行省略，notes_url 缺失为空值行；
    published 不参与签名。独立函数，供测试对拍与 make-manifest.ps1 的构造保持逐字节一致。"""
    lines = ["AttachLink manifest v1"]
    lines.append("version={}".format(m.get("version") if m.get("version") is not None else ""))
    lines.append("url={}".format(m.get("url") if m.get("url") is not None else ""))
    if m.get("size_zip") is not None:
        lines.append("size_zip={}".format(m["size_zip"]))
    lines.append("sha256_zip={}".format(str(m.get("sha256_zip") or "").lower()))
    lines.append("sha256_exe={}".format(str(m.get("sha256_exe") or "").lower()))
    lines.append("notes_url={}".format(m.get("notes_url") if m.get("notes_url") is not None else ""))
    return "\n".join(lines) + "\n"


def verify_manifest_signature(m: dict) -> bool:
    """验签 latest.json：RSASSA-PKCS#1 v1.5 + SHA-256，纯标准库实现（安装器必须零依赖可验签）。
    通过返回 True；任何异常（字段缺失 / base64 非法 / 长度不对 / 签名值 ≥ 模数）都返回 False，不抛出。
    公钥在调用时读取模块级常量，便于单测打桩替换成测试密钥。"""
    try:
        if not isinstance(m, dict):
            return False
        if m.get("sig_alg") != SIG_ALG:
            return False
        sig_b64 = m.get("signature")
        if not isinstance(sig_b64, str):
            return False
        sig = base64.b64decode(sig_b64, validate=True)
        n = RELEASE_KEY_N if isinstance(RELEASE_KEY_N, int) else int(RELEASE_KEY_N, 16)
        e = int(RELEASE_KEY_E)
        k = (n.bit_length() + 7) // 8          # 模数字节数（RSA 2048 → 256）
        if len(sig) != k:
            return False
        s = int.from_bytes(sig, "big")
        if s >= n:
            return False
        em = pow(s, e, n).to_bytes(k, "big")
        digest = hashlib.sha256(manifest_signing_text(m).encode("utf-8")).digest()
        # PS 填充：00 01 FF…FF 00 ‖ DigestInfo(SHA-256) ‖ 摘要；长度全部由常量推导
        tail_len = len(_SIG_DIGESTINFO_PREFIX) + len(digest)
        expected = b"\x00\x01" + b"\xff" * (k - 3 - tail_len) + b"\x00" + _SIG_DIGESTINFO_PREFIX + digest
        return hmac.compare_digest(em, expected)
    except Exception:
        return False


def _verify_manifest_cli(path: str) -> int:
    """python packaging\\setup_update.py --verify-manifest <latest.json>：字段校验 + 验签。
    打印 OK（退出码 0）或失败原因（退出码 1）；供发布 SOP 与 build\\make-manifest.ps1 互操作自检使用。"""
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            m = json.load(f)
    except Exception as e:
        print(f"失败：清单文件无法读取或不是合法 JSON：{e}")
        return 1
    try:
        validate_manifest(m)
    except UpdateError as e:
        print(f"失败：清单字段校验：{e.message}")
        return 1
    if not verify_manifest_signature(m):
        print("失败：清单签名校验失败（不是 AttachLink 发布密钥签发）")
        return 1
    print("OK")
    return 0


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--verify-manifest":
        sys.exit(_verify_manifest_cli(sys.argv[2]))


def _prefix_candidates(mirror=None):
    """下载地址前缀候选（去重保序）：调用方指定 > 上次成功 > 内置镜像列表 > 直连"""
    cands = []
    if mirror:
        cands.append(mirror)
    global _last_good_prefix
    if _last_good_prefix and _last_good_prefix not in cands:
        cands.append(_last_good_prefix)
    for p in MIRROR_PREFIXES:
        if p not in cands:
            cands.append(p)
    if "" not in cands:
        cands.append("")
    return cands


def _monotonic() -> float:
    """单调时钟（独立出来便于单测打桩：总时限用例注入假时钟）"""
    return time.monotonic()


def _parse_manifest_body(body: bytes, relax: bool):
    """把清单响应体解析成可信清单：JSON 解析 → 字段校验 → 发布密钥验签。
    返回 (manifest, reason)：reason 非 None 表示该内容不可用（逐类给出原因）。"""
    try:
        m = json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, ValueError) as e:
        return None, f"内容不是有效的 UTF-8 JSON（{e}）"
    try:
        validate_manifest(m, relax_url_gate=relax)
    except UpdateError as e:
        return None, f"清单字段非法（{e.message}）"
    # 字段合法 ≠ 可信：必须再用内置公钥验签，通过后才信任其中的哈希
    # （本地回环测试模式同样验签，不设旁路——镜像只负责传输，不能伪造清单）
    if not verify_manifest_signature(m):
        return None, ("清单签名校验失败：该地址返回的清单不是 AttachLink 发布密钥签发的"
                      "（镜像被篡改或损坏），已弃用其结果")
    return m, None


def fetch_manifest(mirror=None, log=print, manifest_url=None, current_version=None):
    """获取并校验最新版本清单 latest.json，返回 (清单, 生效的加速前缀)。
    某个前缀取不到内容（网络错误 / 非 200）或内容非法（非 JSON / 字段校验失败 / 清单异常大 /
    最终地址未过校验）都记录原因并换下一个前缀；整轮获取设 30 秒总时限，超时不再尝试剩余前缀；
    全部失败时——若全部是「取不到内容」抛 UpdateError("network")，否则抛 UpdateError("manifest")，
    消息里逐条列出试过的地址与原因。
    current_version 传入时启用「直连复核」：加速镜像（非直连）给出有效签名清单、但其版本不高于
    当前版本（即「说没有新版本」）时，在总时限内再用 GitHub 直连取一次——镜像缓存了旧清单时
    也不至于误报「已是最新」；直连不通（国内常见）则沿用镜像结果并在日志注明可能有缓存延迟。"""
    global _last_good_prefix
    base = manifest_url or MANIFEST_URL
    relax = _is_loopback_url(base)   # 清单来自回环地址 → 本地测试，放宽清单内下载地址门槛
    errors = []
    saw_bad_content = False
    # 回环地址（本地测试）不能拼加速前缀：公共加速站到不了它，只会逐个超时（最长 8 秒 × 前缀数）——只直连
    candidates = [""] if relax else _prefix_candidates(mirror)
    start = _monotonic()
    for prefix in candidates:
        if _monotonic() - start > MANIFEST_TOTAL_TIMEOUT:
            # 总时限已到：不再尝试剩余前缀（每个前缀最长 8 秒，逐个超时会拖到几分钟）
            errors.append(f"总时限已到（{MANIFEST_TOTAL_TIMEOUT} 秒），不再尝试剩余地址")
            log(f"清单获取总时限已到（{MANIFEST_TOTAL_TIMEOUT} 秒），停止尝试剩余地址")
            break
        url = prefix + base
        log(f"获取清单：{url}")
        reason = None
        manifest = None
        try:
            status, _ctype, body = _http_get(url, MANIFEST_TIMEOUT, allow_loopback_http=True,
                                             max_bytes=MAX_MANIFEST_BYTES)
        except UpdateError as e:
            reason = e.message
            if e.code == "manifest":
                # 清单异常大 / 最终地址未过校验：内容本身不可信，与「字段非法」同归类
                saw_bad_content = True
        else:
            if status != 200:
                reason = f"HTTP {status}"
            else:
                manifest, reason = _parse_manifest_body(body, relax)
                if reason is not None:
                    saw_bad_content = True
        if reason is None:
            # 直连复核：加速镜像说「没有新版本」时，向权威来源（GitHub 直连）多问一次，
            # 防止镜像缓存旧清单导致误报「已是最新」。镜像有新版时立即返回（有效签名的新版本
            # 清单本身就是可用的，不必复核）；回环模式只有直连一个候选，天然不会进这里。
            do_recheck = bool(prefix and current_version and parse_version(current_version) is not None
                              and parse_version(manifest.get("version")) is not None
                              and not is_newer(manifest["version"], current_version))
            # 总时限是软预算：发起复核前看一眼剩余时间，复核超时取「剩余预算」与 5 秒的较小值（下限 1 秒）
            remaining = (MANIFEST_TOTAL_TIMEOUT - (_monotonic() - start)) if do_recheck else 0.0
            if do_recheck and remaining > 0:
                log("加速镜像返回的清单版本不高于当前版本，用 GitHub 直连复核一次（防镜像缓存旧清单）…")
                try:
                    d_status, _dc, d_body = _http_get(base, max(1.0, min(MANIFEST_TIMEOUT, 5.0, remaining)),
                                                      allow_loopback_http=True,
                                                      max_bytes=MAX_MANIFEST_BYTES)
                    d_manifest, d_reason = (None, f"HTTP {d_status}") if d_status != 200 \
                        else _parse_manifest_body(d_body, relax)
                except UpdateError as e:
                    d_manifest, d_reason = None, e.message
                if d_manifest is not None and not is_newer(manifest["version"], d_manifest["version"]):
                    # 直连那份不比镜像旧（通常更新）：以直连为准；清空「上次成功的镜像」记录
                    # （后续下载仍按「镜像优先、直连兜底」的顺序——镜像通常比直连快）
                    log("加速镜像可能存在缓存，已用 GitHub 直连复核（以直连结果为准）")
                    _last_good_prefix = ""
                    return d_manifest, ""
                if d_manifest is not None:
                    log("直连清单的版本低于加速镜像，保留加速镜像结果")
                else:
                    log(f"直连复核未成功（{d_reason}），沿用加速镜像结果（经加速镜像取得，可能有缓存延迟）")
            _last_good_prefix = prefix      # 走到这里说明这个镜像前缀取到了有效清单：下次优先
            return manifest, prefix
        errors.append(f"{url}：{reason}")
        log(f"该地址失败，换下一个：{reason}")
    code = "manifest" if saw_bad_content else "network"
    raise UpdateError(code, "获取更新清单失败，尝试过的地址：" + "；".join(errors))


# ---------------- 下载 ----------------

def download_package(manifest, dest_dir, mirror=None, on_progress=None, log=print):
    """下载发布 zip 到 dest_dir 并逐项校验（HTTP 200 / 非 text/html / PK 魔数 / 大小 / SHA-256），
    任一地址不符即删 .part 换下一个前缀，全部失败抛 UpdateError；成功返回 zip 路径。"""
    url = manifest["url"]
    version = manifest["version"]
    os.makedirs(dest_dir, exist_ok=True)
    part = os.path.join(dest_dir, f"AttachLink-{version}.zip.part")
    final = os.path.join(dest_dir, f"AttachLink-{version}.zip")
    # 回环地址（本地测试）不能拼镜像前缀，只直连
    prefixes = [""] if _is_loopback_url(url) else _prefix_candidates(mirror)
    reasons = []
    saw_verify_failure = False
    for prefix in prefixes:
        full = prefix + url
        log(f"下载：{full}")
        _remove_quiet(part)
        try:
            _download_attempt(full, part, manifest, on_progress)
            os.replace(part, final)
            global _last_good_prefix
            _last_good_prefix = prefix
            log(f"下载完成并通过校验：{final}")
            return final
        except UpdateError as e:
            reasons.append(f"{full}：{e.message}")
            if e.code != "network":
                saw_verify_failure = True
        except _VerifyFailure as e:
            reasons.append(f"{full}：{e.reason}")
            saw_verify_failure = True
        finally:
            _remove_quiet(part)
        log(f"该地址失败，换下一个：{reasons[-1]}")
    code = "verify" if saw_verify_failure else "network"
    raise UpdateError(code, "全部下载地址均未通过：" + "；".join(reasons))


def _download_attempt(url, part_path, manifest, on_progress=None):
    """单地址下载尝试：流式写 .part，边下边验（魔数 / 大小上限），收尾验（状态 / 类型 / 大小 / SHA-256）"""
    expected_size = manifest.get("size_zip")
    expected_sha = manifest["sha256_zip"].lower()
    hasher = hashlib.sha256()
    state = {"count": 0, "head": b"", "failed": None}

    def _fail(reason):
        if state["failed"] is None:
            state["failed"] = reason

    def write(block):
        if state["failed"]:
            raise _VerifyFailure(state["failed"])   # 中途已判失败：停止读取，外层负责删 .part
        if len(state["head"]) < 4:
            state["head"] = (state["head"] + block[:4])[:4]
            if len(state["head"]) == 4 and state["head"] != b"PK\x03\x04":
                state["failed"] = "文件不是 zip 包（缺少 PK 魔数）"
                raise _VerifyFailure(state["failed"])
        state["count"] += len(block)
        if state["count"] > MAX_PACKAGE_BYTES:
            state["failed"] = f"文件超过大小上限（{MAX_PACKAGE_BYTES} 字节）"
            raise _VerifyFailure(state["failed"])
        dst.write(block)
        hasher.update(block)

    def on_response(status, content_type, total):
        if status != 200:
            _fail(f"HTTP {status}（应为 200）")
        elif "text/html" in (content_type or "").lower():
            _fail("返回的是网页而不是文件（Content-Type: text/html）")
        elif total > MAX_PACKAGE_BYTES:
            _fail(f"Content-Length 超过大小上限（{total} 字节）")
        if state["failed"]:
            raise _VerifyFailure(state["failed"])

    def report(done, total):
        if on_progress:
            on_progress(done, total if total >= 0 else (expected_size if isinstance(expected_size, int) else -1))

    with open(part_path, "wb") as dst:
        _http_download(url, HTTP_TIMEOUT, write, on_response=on_response, on_chunk=report,
                       allow_loopback_http=True)
    if state["failed"]:
        raise _VerifyFailure(state["failed"])
    if expected_size is not None and state["count"] != expected_size:
        raise _VerifyFailure(f"大小与清单不符（收到 {state['count']} 字节，清单 {expected_size} 字节）")
    actual = hasher.hexdigest()
    if actual != expected_sha:
        raise _VerifyFailure(f"SHA-256 校验不符（清单 {expected_sha[:16]}…，实际 {actual[:16]}…）")


# ---------------- 解压 ----------------

def _entry_parts(name: str):
    return re.split(r"[/\\]", name)


def _is_absolute_entry(name: str) -> bool:
    return name.startswith(("/", "\\")) or bool(re.match(r"^[A-Za-z]:", name))


def extract_installer(zip_path, manifest, dest_dir):
    """从发布 zip 取出顶层唯一的 AttachLinkSetup.exe（绝对路径 / 含 .. 的条目名直接拒绝，防 zip-slip），
    落盘后按清单 SHA-256 复验（不符删文件抛 verify），返回 exe 路径。"""
    expected_sha = manifest["sha256_exe"].lower()
    with zipfile.ZipFile(zip_path) as zf:
        tops = [i for i in zf.infolist()
                if not i.is_dir() and "/" not in i.filename.replace("\\", "/")
                and i.filename.lower() == _INSTALLER_NAME.lower()]
        if len(tops) != 1:
            raise UpdateError("extract",
                              f"压缩包顶层应恰有一个 {_INSTALLER_NAME}，实际找到 {len(tops)} 个")
        info = tops[0]
        if any(p == ".." for p in _entry_parts(info.filename)) or _is_absolute_entry(info.filename):
            raise UpdateError("extract", f"压缩包条目路径可疑（含 .. 或绝对路径），已拒绝：{info.filename}")
        # 全部检查通过后才建目标目录 / 落盘（zip-slip 情况不产生任何文件）
        os.makedirs(dest_dir, exist_ok=True)
        target = os.path.join(dest_dir, _INSTALLER_NAME)   # 只用条目 basename，固定文件名
        hasher = hashlib.sha256()
        count = 0
        try:
            with zf.open(info) as src, open(target, "wb") as dst:
                while True:
                    block = src.read(256 * 1024)
                    if not block:
                        break
                    count += len(block)
                    if count > MAX_PACKAGE_BYTES:
                        raise _VerifyFailure(f"条目超过大小上限（{MAX_PACKAGE_BYTES} 字节）")
                    dst.write(block)
                    hasher.update(block)
        except _VerifyFailure as e:
            _remove_quiet(target)
            raise UpdateError("verify", f"解压中止：{e.reason}") from e
        actual = hasher.hexdigest()
        if actual != expected_sha:
            _remove_quiet(target)
            raise UpdateError("verify",
                              f"安装器 SHA-256 校验不符（清单 {expected_sha[:16]}…，实际 {actual[:16]}…）")
    return target


# ---------------- 交接 ----------------

def build_handoff_args(new_exe, old_exe, old_pid, action="/install-all"):
    """新安装器命令行：执行安装动作，并携带 --updated-from / --wait-pid 交接参数"""
    return [new_exe, action, "--updated-from", old_exe, "--wait-pid", str(old_pid)]


def launch_handoff(args):
    """启动新安装器进程（不走 shell），返回其 PID"""
    proc = subprocess.Popen(args, close_fds=True)
    return proc.pid


def _running_exe() -> str:
    """当前安装器 exe 路径（冻结运行=exe 本身；源码运行=python.exe）。独立出来便于单测打桩"""
    return sys.executable


def _wait_process_exit(pid, timeout) -> bool:
    """轮询等待指定进程退出（OpenProcess(SYNCHRONIZE) + WaitForSingleObject），返回是否确认退出；
    进程打不开（已退出 / 无权限）按已退出处理，替换步骤自会暴露真实问题。"""
    import ctypes
    SYNCHRONIZE = 0x00100000
    WAIT_OBJECT_0 = 0x00000000
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(SYNCHRONIZE, False, int(pid))
    if not handle:
        return True
    try:
        deadline = time.time() + max(0.0, float(timeout))
        while time.time() < deadline:
            if kernel32.WaitForSingleObject(handle, 250) == WAIT_OBJECT_0:
                return True
        return False
    finally:
        kernel32.CloseHandle(handle)


def _copy_self_ready(running: str, target_dir: str):
    """把本安装器复制成 target_dir 下「内容已校验」的临时文件：先复制到同目录唯一临时文件
    （AttachLinkSetup.exe.new-<随机>），再算 SHA-256 与本安装器自身比对。返回 (临时路径, 错误信息)：
    失败时删掉临时文件并返回 (None, 原因)——校验通过之前不碰任何旧文件。
    备份名同样用随机后缀（AttachLinkSetup.exe.bak-<随机>），避免与上次残留的 .bak 冲突。"""
    tmp = os.path.join(target_dir, f"{_INSTALLER_NAME}.new-{uuid.uuid4().hex}")
    try:
        shutil.copyfile(running, tmp)
        if _sha256_file(tmp) != _sha256_file(running):
            raise OSError("复制后内容校验不一致（SHA-256 不符）")
    except OSError as e:
        _remove_quiet(tmp)
        return None, str(e)
    return tmp, ""


def replace_old_exe(old_exe, wait_pid=None, timeout=15):
    """交接自检：等旧安装器进程退出后，把旧 exe 替换为本安装器的副本。
    安全顺序——先把本安装器复制到旧文件同目录的临时文件并校验 SHA-256（内容就绪并验证通过之前
    不碰旧文件），再把旧文件改名为随机后缀备份，最后把临时文件就位；任一步失败都如实报告：
    替换失败时尝试回滚，回滚失败则保留备份路径并要求用户改回原名。旧文件不存在时：父目录可写
    则按新版本在原位置重建（recreated），否则如实报缺失（missing），绝不谎称「旧安装器文件已更新」。
    返回结构化结果 {"ok", "status", "message", "backup"}：
    status ∈ replaced（已替换）/ recreated（原位置已重建）/ same（旧安装器即本安装器）/
    missing（旧文件不存在且原位置不可写，ok=False）/ failed（失败，ok=False，backup 可能非空）。"""
    notes = []
    if wait_pid is not None:
        try:
            pid = int(wait_pid)
        except (TypeError, ValueError):
            pid = 0
            notes.append(f"--wait-pid 不是数字（{wait_pid!r}），跳过等待")
        if pid > 0:
            if _wait_process_exit(pid, timeout):
                notes.append(f"旧安装器进程（PID {pid}）已退出")
            else:
                return {"ok": False, "status": "failed",
                        "message": "；".join(notes + [f"等待旧安装器进程（PID {pid}）退出超时（{timeout} 秒），本次未替换"]),
                        "backup": None}
    running = _running_exe()
    if not old_exe:
        return {"ok": False, "status": "missing",
                "message": "；".join(notes + ["未提供旧安装器路径（--updated-from 为空），未替换"]), "backup": None}
    parent = os.path.dirname(os.path.abspath(old_exe))
    if not os.path.exists(old_exe):
        # 旧文件已不存在：父目录存在且可写 → 按新版本重建；否则如实报缺失（ok=False），
        # 由 handoff_summary_line 据 status 给出「请把它放到固定位置」的提示
        if not os.path.isdir(parent):
            return {"ok": False, "status": "missing",
                    "message": "；".join(notes + [f"旧安装器文件已不存在（{old_exe}），且原位置的目录不存在，未替换"]),
                    "backup": None}
        tmp, err = _copy_self_ready(running, parent)
        if tmp is None:
            return {"ok": False, "status": "missing",
                    "message": "；".join(notes + [f"旧安装器文件已不存在，且原位置不可写（{err}），未替换"]),
                    "backup": None}
        try:
            os.replace(tmp, old_exe)
        except OSError as e:
            _remove_quiet(tmp)
            return {"ok": False, "status": "missing",
                    "message": "；".join(notes + [f"旧安装器文件已不存在，且原位置不可写（{e}），未替换"]),
                    "backup": None}
        notes.append("原位置的安装器此前已不存在，已按新版本重建")
        return {"ok": True, "status": "recreated", "message": "；".join(notes), "backup": None}
    try:
        same = os.path.exists(running) and os.path.samefile(old_exe, running)
    except OSError:
        same = False
    if same:
        notes.append("旧安装器与本安装器是同一文件，无需替换")
        return {"ok": True, "status": "same", "message": "；".join(notes), "backup": None}
    # 1) 本安装器复制到同目录临时文件并 SHA-256 校验（只比大小的时代已结束：等长错误内容也能识别）
    tmp, err = _copy_self_ready(running, parent)
    if tmp is None:
        return {"ok": False, "status": "failed",
                "message": "；".join(notes + [f"复制本安装器到临时文件失败（{err}），旧安装器未被动过"]),
                "backup": None}
    # 2) 旧文件改名为随机后缀备份（留底，防中途失败无回滚余地）
    backup = os.path.join(parent, f"{_INSTALLER_NAME}.bak-{uuid.uuid4().hex}")
    try:
        os.replace(old_exe, backup)
    except OSError as e:
        _remove_quiet(tmp)
        return {"ok": False, "status": "failed",
                "message": "；".join(notes + [f"旧安装器改名失败（可能被占用）：{e}，旧文件未变"]),
                "backup": None}
    # 3) 临时文件就位；失败则回滚（把备份改回原名）
    try:
        os.replace(tmp, old_exe)
    except OSError as e:
        _remove_quiet(tmp)
        try:
            os.replace(backup, old_exe)
        except OSError as e2:
            return {"ok": False, "status": "failed",
                    "message": "；".join(notes + [
                        f"替换旧安装器失败（{e}），且回滚也失败（{e2}）："
                        f"原位置目前没有安装器；旧文件的备份保留在：{backup}——请把它改回原文件名"]),
                    "backup": backup}
        return {"ok": False, "status": "failed",
                "message": "；".join(notes + [f"替换旧安装器失败（已回滚，旧文件未变）：{e}"]),
                "backup": None}
    try:
        os.remove(backup)
    except OSError:
        notes.append(f"备份删除失败，已保留：{backup}")
    notes.append("旧安装器已替换为新版本")
    return {"ok": True, "status": "replaced", "message": "；".join(notes), "backup": None}


# ---------------- 主流程 ----------------

def _remove_quiet(path: str) -> None:
    """删除文件（尽力而为：不存在 / 被占用时静默跳过）"""
    try:
        os.remove(path)
    except OSError:
        pass


_TMP_MAX_AGE_HOURS = 24   # 更新临时子目录的保留时长：启动清理只删超过此时长的旧子目录


def _update_tmp_root() -> str:
    """更新临时目录根：%TEMP%\\AttachLinkSetup_update"""
    d = os.path.join(tempfile.gettempdir(), _TMP_DIR_NAME)
    os.makedirs(d, exist_ok=True)
    return d


def _update_tmp_dir() -> str:
    """本次更新的专用子目录：根目录下按次生成 GUID 目录——两个安装器同时更新各用各的目录，
    失败时只清理自己的子目录，成功交接时保留（新安装器正从那里运行）。
    （旧实现多实例共用单一目录、启动即整体清空，两个安装器同时更新会互删文件。）"""
    d = os.path.join(_update_tmp_root(), uuid.uuid4().hex)
    os.makedirs(d, exist_ok=True)
    return d


def _clean_stale_tmp_dirs(root: str, max_age_hours: float | None = None, now: float | None = None) -> int:
    """只清理 root 下修改时间超过 max_age_hours（默认 24 小时）的子目录，返回清理个数；
    正在使用的新子目录与文件项一律不动。now 可注入（单测用）。"""
    if max_age_hours is None:
        max_age_hours = _TMP_MAX_AGE_HOURS
    if now is None:
        now = time.time()
    try:
        names = os.listdir(root)
    except OSError:
        return 0
    removed = 0
    for name in names:
        p = os.path.join(root, name)
        try:
            if not os.path.isdir(p):
                continue
            if now - os.path.getmtime(p) > max_age_hours * 3600:
                shutil.rmtree(p, ignore_errors=True)
                removed += 1
        except OSError:
            continue
    return removed


def _clean_tmp_dir(d: str) -> None:
    """清理更新临时目录（尽力而为，失败忽略）"""
    shutil.rmtree(d, ignore_errors=True)


def _sha256_file(path: str) -> str:
    hasher = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(256 * 1024), b""):
            hasher.update(block)
    return hasher.hexdigest()


def run_update(dry_run=False, mirror=None, manifest_url=None, action="/install-all",
               out=print, progress=None):
    """一键更新主流程，返回给用户看的汇总文本（不抛异常，失败以「失败：…」开头）：
    清单 → 版本比较（已是最新直接返回）→ 下载 → 校验 → 解压 → dry_run 到此为止 → 交接新安装器。"""
    tmp_dir = _update_tmp_dir()                        # 本次运行专属 GUID 子目录（并发更新互不干扰）
    _clean_stale_tmp_dirs(os.path.dirname(tmp_dir))    # 只清超过 24 小时的旧子目录，不动正在使用的
    try:
        out("检查更新：正在获取最新版本信息…")
        current = _current_version()
        manifest, worked = fetch_manifest(mirror=mirror, manifest_url=manifest_url, current_version=current,
                                          log=lambda s: out(s))
        if worked:
            out(f"获取清单成功（经 {worked or '直连'}）")
        latest = manifest["version"]
        if parse_version(current) is not None and not is_newer(latest, current):
            msg = f"当前已是最新版本 v{current}"
            if worked:
                # 经加速镜像取得（直连不可达/未用）且镜像说没有新版：注明可能有缓存延迟
                msg += "（据加速镜像，新版本发布后可能有缓存延迟；如刚发布新版，可稍后再查或到 Releases 页面手动下载）"
            out(msg)
            return msg
        if parse_version(current) is None:
            out("当前版本号不可读，按有新版本处理")
        out(f"发现新版本 v{latest}（当前 v{current}），开始下载…")
        zip_path = download_package(manifest, tmp_dir, mirror=mirror,
                                    on_progress=progress, log=lambda s: out(s))
        out("下载完成且校验通过，正在解压安装器…")
        exe_path = extract_installer(zip_path, manifest, tmp_dir)
        exe_sha = _sha256_file(exe_path)
        if dry_run:
            msg = (f"试运行完成：已下载并校验 v{latest}，安装器位于 {exe_path}"
                   f"（SHA-256 前16位 {exe_sha[:16]}）")
            out(msg)
            return msg
        if not getattr(sys, "frozen", False):
            _clean_tmp_dir(tmp_dir)
            msg = "失败：交接安装仅支持发布版安装器（exe 直接运行）；源码运行请加 --update-dry-run 验证下载链路"
            out(msg)
            return msg
        old_exe = os.path.abspath(sys.executable)
        args = build_handoff_args(exe_path, old_exe, os.getpid(), action=action)
        out(f"启动新版本安装器 v{latest} 交接安装（本窗口即将关闭）…")
        new_pid = launch_handoff(args)
        _remove_quiet(zip_path)   # 交接启动成功：本次下载的 zip 不再保留（解压出的新安装器正在运行，必须留下）
        msg = f"已启动新版本安装器（PID {new_pid}）"
        out(msg)
        return msg
    except UpdateError as e:
        _clean_tmp_dir(tmp_dir)
        msg = f"失败：{e.message}（{e.code}）{_SUGGESTIONS.get(e.code, '')}"
        out(msg)
        return msg
    except Exception as e:  # noqa: BLE001  兜底：任何未预期异常都转成失败文本，不让界面线程崩
        _clean_tmp_dir(tmp_dir)
        msg = f"失败：更新过程中出现异常：{e}"
        out(msg)
        return msg

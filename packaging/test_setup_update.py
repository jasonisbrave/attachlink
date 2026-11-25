# setup_update.py 无头单测：不访问外网，HTTP 一律打桩（mock urllib.request.urlopen），文件操作走临时目录。
# 覆盖：版本解析与比较、清单校验逐类非法样例、fetch_manifest 逐前缀回退（内容非法也换）、_http_get 打桩四态、
# download_package 逐镜像回退、extract_installer（正常 / 无 exe / 多 exe / zip-slip / 哈希不符）、build_handoff_args、
# replace_old_exe（结构化字典：replaced / recreated / same / missing / failed，回滚失败如实报告备份路径）、
# 中途断流换来源（B3）、镜像缓存直连复核（B4）、run_update 汇总文本与交接成功后删 zip（B5）。
# 用法：python packaging\test_setup_update.py   期望末行 RESULT: PASS
import base64
import hashlib
import http.client
import importlib.util
import io
import json
import math
import os
import secrets
import subprocess
import sys
import tempfile
import urllib.error
import zipfile
from unittest import mock

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)  # 让 setup_update 内部的延迟 import setup_actions 可用

spec = importlib.util.spec_from_file_location("setup_update", os.path.join(_HERE, "setup_update.py"))
su = importlib.util.module_from_spec(spec)
sys.modules["setup_update"] = su  # 注册：setup_actions 的 import setup_update 复用同一模块对象（打桩才生效）
spec.loader.exec_module(su)

ok = True
def check(cond, label):
    global ok
    print(("pass: " if cond else "FAIL: ") + label)
    if not cond: ok = False

def expect_update_error(fn, code, label, contains=None):
    try:
        fn()
        check(False, label + "（未抛出）")
    except su.UpdateError as e:
        good = e.code == code and (contains is None or contains in e.message)
        check(good, label + ("" if good else f"（实际 code={e.code} message={e.message}）"))


# ---------------- 0) 测试用 RSA 密钥对（1024 位，纯 Python：secrets + Miller-Rabin） ----------------
# 清单签名规格要求「测试侧实现 PKCS#1 v1.5 签名」：生成自己的密钥对，打桩替换 su.RELEASE_KEY_N/E，
# 绝不读取真实发布私钥。正文 / 填充格式与 setup_update.verify_manifest_signature 的期望逐字节一致。
def _is_probable_prime(x, rounds=40):
    if x < 2:
        return False
    for sp in (2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37):
        if x % sp == 0:
            return x == sp
    d = x - 1
    r = 0
    while d % 2 == 0:
        d //= 2
        r += 1
    for _ in range(rounds):
        a = secrets.randbelow(x - 3) + 2
        v = pow(a, d, x)
        if v in (1, x - 1):
            continue
        for _ in range(r - 1):
            v = pow(v, 2, x)
            if v == x - 1:
                break
        else:
            return False
    return True

def _gen_prime(bits):
    while True:
        c = secrets.randbits(bits) | (1 << (bits - 1)) | 1
        if _is_probable_prime(c):
            return c

def _make_keypair():
    while True:
        p, q = _gen_prime(512), _gen_prime(512)
        if p == q:
            continue
        lam = math.lcm(p - 1, q - 1)
        if math.gcd(65537, lam) != 1:
            continue
        return p * q, 65537, pow(65537, -1, lam)

_TEST_N, _TEST_E, _TEST_D = _make_keypair()
_TEST_N2, _TEST_E2, _TEST_D2 = _make_keypair()   # 第二对：「错密钥签名必须失败」用例

def pkcs1v15_sign(data: bytes, n, d) -> bytes:
    """测试侧 PKCS#1 v1.5 + SHA-256 签名（与发布侧 SignData(bytes, SHA256-OID) 同一格式）"""
    k = (n.bit_length() + 7) // 8
    digest = hashlib.sha256(data).digest()
    prefix = bytes.fromhex("3031300d060960864801650304020105000420")
    tail = prefix + digest
    em = b"\x00\x01" + b"\xff" * (k - 3 - len(tail)) + b"\x00" + tail
    return pow(int.from_bytes(em, "big"), d, n).to_bytes(k, "big")

def sign_manifest(m) -> str:
    """用测试私钥给清单签名（返回 base64）——签名正文按 su.manifest_signing_text 生成"""
    return base64.b64encode(pkcs1v15_sign(su.manifest_signing_text(m).encode("utf-8"), _TEST_N, _TEST_D)).decode("ascii")

def sign_manifest_wrong_key(m) -> str:
    """用第二对测试密钥签名（模拟「另一个发布者」的签名）"""
    return base64.b64encode(pkcs1v15_sign(su.manifest_signing_text(m).encode("utf-8"), _TEST_N2, _TEST_D2)).decode("ascii")

# 打桩：验签改用测试公钥（模块级常量调用时读取，这里直接替换）
su.RELEASE_KEY_N = format(_TEST_N, "x")
su.RELEASE_KEY_E = _TEST_E

# ---------------- 1) 版本解析与比较 ----------------
check(su.parse_version("1.4.5") == (1, 4, 5), "parse_version：1.4.5 → (1,4,5)")
check(su.parse_version("1.10.0") == (1, 10, 0), "parse_version：逐段按整数比（1.10 > 1.9）")
check(su.parse_version(" 1.4 ") == (1, 4), "parse_version：容忍首尾空白、两段版本")
check(su.parse_version("") is None and su.parse_version("abc") is None
      and su.parse_version("1..4") is None and su.parse_version("1.4.5-beta") is None
      and su.parse_version(None) is None and su.parse_version(1.45) is None,
      "parse_version：空串/非数字/空段/后缀/非字符串 → None")
check(su.is_newer("1.4.5", "1.4.4") is True, "is_newer：新 > 旧 → True")
check(su.is_newer("1.4.4", "1.4.4") is False, "is_newer：相等 → False（严格大于）")
check(su.is_newer("1.4.3", "1.4.4") is False, "is_newer：旧 < 新 → False")
check(su.is_newer("1.4.5", "未知版本") is False, "is_newer：任一无法解析 → False")

# ---------------- 2) 清单校验（每类非法各一例） ----------------
GOOD_URL = "https://github.com/jasonisbrave/attachlink/releases/download/v1.4.5/AttachLink-1.4.5.zip"
SHA = "a" * 64
def manifest(**over):
    m = {"version": "1.4.5", "published": "2026-11-25", "url": GOOD_URL,
         "size_zip": 123, "sha256_zip": SHA, "sha256_exe": SHA,
         "notes_url": "https://github.com/jasonisbrave/attachlink/releases/tag/v1.4.5"}
    m.update(over)
    m = {k: v for k, v in m.items() if v is not None}
    m["sig_alg"] = su.SIG_ALG
    m["signature"] = sign_manifest(m)   # 合法清单一律带测试密钥签名（sig 字段不参与正文）
    return m

def tamper_signed(m, **over):
    """对已签名的清单原地篡改字段（签名保持原样）——用于「篡改必须验签失败」用例"""
    m.update(over)
    return m

su.validate_manifest(manifest())  # 合法清单静默通过
check(True, "validate_manifest：合法清单通过")
expect_update_error(lambda: su.validate_manifest("not a dict"), "manifest", "清单校验：非 JSON 对象")
expect_update_error(lambda: su.validate_manifest(manifest(version="abc")), "manifest", "清单校验：version 非法")
expect_update_error(lambda: su.validate_manifest(manifest(url="http://github.com/jasonisbrave/attachlink/releases/download/v1.4.5/AttachLink-1.4.5.zip")),
                    "manifest", "清单校验：url 非 https")
expect_update_error(lambda: su.validate_manifest(manifest(url="https://evil.example.com/jasonisbrave/attachlink/releases/download/v1.4.5/AttachLink-1.4.5.zip")),
                    "manifest", "清单校验：url 主机不对", contains="github.com")
expect_update_error(lambda: su.validate_manifest(manifest(url="https://github.com/other/repo/releases/download/v1.4.5/AttachLink-1.4.5.zip")),
                    "manifest", "清单校验：url 路径前缀不对")
expect_update_error(lambda: su.validate_manifest(manifest(url="https://github.com/jasonisbrave/attachlink/releases/download/v1.4.5/AttachLink-1.4.5.exe")),
                    "manifest", "清单校验：url 不以 .zip 结尾")
expect_update_error(lambda: su.validate_manifest(manifest(sha256_zip="a" * 63)), "manifest", "清单校验：哈希 63 位")
expect_update_error(lambda: su.validate_manifest(manifest(sha256_exe="z" * 64)), "manifest", "清单校验：哈希非十六进制")
expect_update_error(lambda: su.validate_manifest(manifest(size_zip=-1)), "manifest", "清单校验：size_zip 为负")
expect_update_error(lambda: su.validate_manifest(manifest(size_zip=True)), "manifest", "清单校验：size_zip 为布尔")
expect_update_error(lambda: su.validate_manifest(manifest(notes_url="http://github.com/x")), "manifest", "清单校验：notes_url 非 https")
expect_update_error(lambda: su.validate_manifest(manifest(url=None)), "manifest", "清单校验：缺 url")

# 回环放宽：本地测试清单允许 http://127.0.0.1，但不放行外网 http / 哈希错误
su.validate_manifest(manifest(url="http://127.0.0.1:8765/AttachLink-1.4.5.zip"), relax_url_gate=True)
check(True, "validate_manifest：relax 下回环 http 清单通过")
expect_update_error(lambda: su.validate_manifest(manifest(url="http://example.com/x.zip"), relax_url_gate=True),
                    "manifest", "清单校验：relax 只放行回环，不放行外网 http")

# ---------------- 3) _http_get（打桩 urlopen：超时 / 404 / HTML / 正常 / 协议闸） ----------------
class FakeResp:
    def __init__(self, body=b"", status=200, ctype="application/json", headers=None):
        self._buf = io.BytesIO(body)
        self.status = status
        self.headers = headers if headers is not None else {"Content-Type": ctype}
    def read(self, n=-1):
        return self._buf.read(n)
    def getcode(self):
        return self.status
    def __enter__(self):
        return self
    def __exit__(self, *exc):
        return False

def fake_urlopen(responses, calls=None, record_headers=None):
    """按 URL 返回预设响应；'*' 兜底；预设值为可调用对象时逐次调用（每次拿全新响应），
    为 Exception 时抛出"""
    def _fake(request, timeout=None):
        url = request.full_url
        if calls is not None:
            calls.append(url)
        if record_headers is not None:
            record_headers.append(dict(request.header_items()))
        item = responses.get(url, responses.get("*"))
        if isinstance(item, Exception):
            raise item
        if callable(item):
            item = item()
        return item
    return _fake

with mock.patch("urllib.request.urlopen", fake_urlopen({"*": TimeoutError()})):
    expect_update_error(lambda: su._http_get("https://github.com/x.json", 1), "network",
                        "_http_get：超时 → network", contains="请求失败")
http404 = urllib.error.HTTPError("https://github.com/x.json", 404, "Not Found", None, io.BytesIO(b"nf"))
with mock.patch("urllib.request.urlopen", fake_urlopen({"*": http404})):
    expect_update_error(lambda: su._http_get("https://github.com/x.json", 1), "network",
                        "_http_get：404 → network", contains="404")
with mock.patch("urllib.request.urlopen", fake_urlopen({"*": FakeResp(b"<html>err</html>", ctype="text/html; charset=utf-8")})):
    status, ctype, body = su._http_get("https://github.com/x.json", 1)
    check(status == 200 and ctype.startswith("text/html") and body == b"<html>err</html>",
          "_http_get：返回 HTML 时如实透传（由调用方按类型拒绝）")
chunks_seen = []
with mock.patch("urllib.request.urlopen", fake_urlopen({"*": FakeResp(b'{"version":"1.4.5"}')})):
    status, ctype, body = su._http_get("https://github.com/x.json", 1, on_chunk=chunks_seen.append)
    check(status == 200 and ctype == "application/json" and body == b'{"version":"1.4.5"}',
          "_http_get：正常返回（状态 / 类型 / 响应体）")
    check(sum(chunks_seen) == len(body) and len(chunks_seen) >= 1, "_http_get：on_chunk 收到全部字节")
seen_headers = []
with mock.patch("urllib.request.urlopen", fake_urlopen({"*": lambda: FakeResp(b"ok")}, record_headers=seen_headers)):
    su._http_get("https://github.com/x.json", 1)
ua = seen_headers[0].get("User-agent", "")
check(isinstance(ua, str) and ua.startswith("AttachLinkSetup/"),
      f"_http_get：User-Agent 为 AttachLinkSetup/<版本>（实际 {ua!r}）")
expect_update_error(lambda: su._http_get("http://github.com/x", 1), "network", "_http_get：http 未放行 → 拒绝")
with mock.patch("urllib.request.urlopen", fake_urlopen({"*": lambda: FakeResp(b"loop")})):
    check(su._http_get("http://127.0.0.1:9/x", 1, allow_loopback_http=True)[2] == b"loop",
          "_http_get：放行后回环 http 可用")
    check(su._http_get("http://localhost:9/x", 1, allow_loopback_http=True)[2] == b"loop",
          "_http_get：放行后 localhost 可用")
    expect_update_error(lambda: su._http_get("http://example.com/x", 1, allow_loopback_http=True),
                        "network", "_http_get：放行只限回环主机，外网 http 仍拒绝")

# ---------------- 2b) fetch_manifest：内容非法也换前缀，全失败按原因分类 ----------------
FP0, FP1, FP2 = su.MIRROR_PREFIXES
VALID_MJSON = json.dumps(manifest(version="1.4.9", size_zip=10)).encode("utf-8")   # 测试密钥签名
su._last_good_prefix = ""
with mock.patch("urllib.request.urlopen",
                fake_urlopen({FP0 + su.MANIFEST_URL: FakeResp(b"<html>mirror 404 page</html>", ctype="text/html; charset=utf-8"),
                              FP1 + su.MANIFEST_URL: FakeResp(VALID_MJSON, ctype="application/json")})):
    m, worked = su.fetch_manifest(log=lambda s: None)
check(m["version"] == "1.4.9" and worked == FP1,
      "fetch_manifest：首个前缀 200+网页 → 换前缀成功，生效前缀是第二个")
check(su._last_good_prefix == FP1, "fetch_manifest：成功后记录上次成功前缀")

bad_all = {p + su.MANIFEST_URL: (lambda: FakeResp(b"<html>err</html>", ctype="text/html")) for p in (FP0, FP1, FP2)}
bad_all[su.MANIFEST_URL] = lambda: FakeResp(b"<html>err</html>", ctype="text/html")
try:
    with mock.patch("urllib.request.urlopen", fake_urlopen(bad_all)):
        su.fetch_manifest(log=lambda s: None)
    check(False, "fetch_manifest：全部前缀内容非法应抛错")
except su.UpdateError as e:
    check(e.code == "manifest" and "gh-proxy.com" in e.message and "https://github.com" in e.message,
          f"fetch_manifest：全部内容非法 → manifest 并逐条列地址（实际 {e.code}）")

net_all = {u: urllib.error.URLError("conn refused") for u in
           [p + su.MANIFEST_URL for p in (FP0, FP1, FP2)] + [su.MANIFEST_URL]}
try:
    with mock.patch("urllib.request.urlopen", fake_urlopen(net_all)):
        su.fetch_manifest(log=lambda s: None)
    check(False, "fetch_manifest：全部前缀网络失败应抛错")
except su.UpdateError as e:
    check(e.code == "network", f"fetch_manifest：全部网络失败 → network（实际 {e.code}）")
su._last_good_prefix = ""  # 还原，避免影响后续用例的前缀顺序

# ---------------- 2a) NB-2：清单大小上限 64 KB（超过即判 manifest「清单异常大」） ----------------
loop_base = "http://127.0.0.1:8765/latest.json"
big_body = b'{"version":"1.4.9", ' + b' ' * (su.MAX_MANIFEST_BYTES + 100) + b'}'
with mock.patch("urllib.request.urlopen", fake_urlopen({loop_base: lambda: FakeResp(big_body, ctype="application/json")})):
    expect_update_error(lambda: su.fetch_manifest(manifest_url=loop_base, log=lambda s: None),
                        "manifest", "fetch_manifest：清单超过 64 KB → manifest「清单异常大」", contains="清单异常大")

# ---------------- 2b2) NB-2：整轮清单获取 30 秒总时限，超时不再请求剩余前缀 ----------------
budget_calls = []
budget_resp = {p + su.MANIFEST_URL: urllib.error.URLError("slow") for p in (FP0, FP1, FP2)}
budget_resp[su.MANIFEST_URL] = urllib.error.URLError("slow")
mono = iter([0.0, 0.0, 999.0])   # start=0.0；第 1 个前缀检查后 0-0 不超时；第 2 次检查 999-0 > 30 → 停止
try:
    with mock.patch("urllib.request.urlopen", fake_urlopen(budget_resp, budget_calls)), \
         mock.patch.object(su, "_monotonic", lambda: next(mono)):
        su.fetch_manifest(log=lambda s: None)
    check(False, "fetch_manifest：总时限到后应抛错")
except su.UpdateError as e:
    check("总时限已到" in e.message, f"fetch_manifest：总时限 → 错误消息说明「总时限已到」（实际 {e.message}）")
check(len(budget_calls) == 1, f"fetch_manifest：总时限到 → 不再请求剩余前缀（实际请求 {len(budget_calls)} 次）")

# ---------------- 2b3) NB-3：最终跳转地址校验（geturl 复查，被降级/跳走 → 拒绝且不读内容） ----------------
class RedirectResp(FakeResp):
    """提供 geturl() 的打桩响应：模拟 urllib 跟随重定向后的最终地址"""
    def __init__(self, body, final_url, **kw):
        super().__init__(body, **kw)
        self._final = final_url
    def geturl(self):
        return self._final

# 最终地址被降级到 http（非回环）→ 拒绝
with mock.patch("urllib.request.urlopen",
                fake_urlopen({loop_base: lambda: RedirectResp(VALID_MJSON, "http://github.com/" + su.REPO + "/evil.zip")})):
    expect_update_error(lambda: su.fetch_manifest(manifest_url=loop_base, log=lambda s: None),
                        "manifest", "_check_final_url：http 降级 → 拒绝", contains="最终地址未通过校验")
# 最终地址跳到白名单外主机 → 拒绝
with mock.patch("urllib.request.urlopen",
                fake_urlopen({loop_base: lambda: RedirectResp(VALID_MJSON, "https://evil.example.com/x.zip")})):
    expect_update_error(lambda: su.fetch_manifest(manifest_url=loop_base, log=lambda s: None),
                        "manifest", "_check_final_url：白名单外主机 → 拒绝", contains="evil.example.com")
# 最终地址合法（github.com https）→ 正常通过
github_final = su.MANIFEST_URL
with mock.patch("urllib.request.urlopen",
                fake_urlopen({loop_base: lambda: RedirectResp(VALID_MJSON, github_final)})):
    m, worked = su.fetch_manifest(manifest_url=loop_base, log=lambda s: None)
check(m["version"] == "1.4.9", "_check_final_url：最终地址 github.com https → 放行")

# ---------------- 2c) 回环清单只直连；User-Agent 仅 ASCII ----------------
loop_base = "http://127.0.0.1:8765/latest.json"
loop_calls = []
with mock.patch("urllib.request.urlopen",
                fake_urlopen({loop_base: FakeResp(VALID_MJSON, ctype="application/json")}, calls=loop_calls)):
    m, worked = su.fetch_manifest(manifest_url=loop_base, log=lambda s: None)
check(m["version"] == "1.4.9" and worked == "" and loop_calls == [loop_base],
      f"fetch_manifest：回环清单地址只直连、不拼加速前缀（实际请求 {loop_calls}）")
with mock.patch.object(su, "_current_version", lambda: "未知版本"):
    check(su._user_agent() == "AttachLinkSetup/unknown", "_user_agent：版本号非法时用纯 ASCII 的 unknown")
with mock.patch.object(su, "_current_version", lambda: "1.4.5"):
    check(su._user_agent() == "AttachLinkSetup/1.4.5", "_user_agent：版本号合法时带版本号")

# ---------------- 2d) 清单签名：正文格式、验签全分支、fetch_manifest 回退、--verify-manifest 入口 ----------------
# 正文逐字节格式：字段顺序 / 行序 / 行尾带 "\n" / published 不参与
EXPECT_TEXT = ("AttachLink manifest v1\n"
               "version=1.4.5\n"
               "url=" + GOOD_URL + "\n"
               "size_zip=123\n"
               "sha256_zip=" + SHA + "\n"
               "sha256_exe=" + SHA + "\n"
               "notes_url=https://github.com/jasonisbrave/attachlink/releases/tag/v1.4.5\n")
check(su.manifest_signing_text(manifest()) == EXPECT_TEXT,
      "manifest_signing_text：字段顺序 / 行序 / 行尾换行逐字节一致（published 不参与签名）")
_nonotes = manifest(notes_url=None, size_zip=None)
check("size_zip=" not in su.manifest_signing_text(_nonotes)
      and su.manifest_signing_text(_nonotes).endswith("\nnotes_url=\n"),
      "manifest_signing_text：size_zip 缺省整行省略、notes_url 缺省为空值行")

# 验签：测试私钥签出的签名必须通过自己的验签（即 DigestInfo/填充字节与发布侧互操作）
check(su.verify_manifest_signature(manifest()), "verify_manifest_signature：测试密钥签出的清单通过")
_k = (su.RELEASE_KEY_N if isinstance(su.RELEASE_KEY_N, int) else int(su.RELEASE_KEY_N, 16)).bit_length()
_k = (_k + 7) // 8
for _f, _v in [("version", "1.4.6"), ("url", GOOD_URL + "x"), ("size_zip", 124),
               ("sha256_zip", "b" * 64), ("sha256_exe", "c" * 64),
               ("notes_url", "https://github.com/jasonisbrave/attachlink/releases/tag/v9.9.9")]:
    check(not su.verify_manifest_signature(tamper_signed(manifest(), **{_f: _v})),
          f"verify_manifest_signature：篡改 {_f} 后签名失效 → False")
check(not su.verify_manifest_signature(tamper_signed(manifest(), sig_alg="RS256-SHA256-PKCS1v15:v2")),
      "verify_manifest_signature：sig_alg 不匹配 → False")
_no_sig = manifest(); _no_sig.pop("signature")
check(not su.verify_manifest_signature(_no_sig), "verify_manifest_signature：缺 signature → False")
check(not su.verify_manifest_signature(tamper_signed(manifest(), signature="!!!not-base64!!!")),
      "verify_manifest_signature：signature 非合法 base64 → False")
check(not su.verify_manifest_signature(tamper_signed(manifest(), signature=base64.b64encode(b"\x00" * (_k - 1)).decode())),
      "verify_manifest_signature：签名长度 ≠ 模数字节数 → False")
check(not su.verify_manifest_signature(tamper_signed(manifest(), signature=base64.b64encode(b"\xff" * _k).decode())),
      "verify_manifest_signature：签名值 ≥ 模数 → False")
check(not su.verify_manifest_signature(tamper_signed(manifest(), signature=sign_manifest_wrong_key(manifest()))),
      "verify_manifest_signature：另一对密钥签的名 → False")
check(not su.verify_manifest_signature("not a dict") and not su.verify_manifest_signature(None),
      "verify_manifest_signature：非 dict 入参 → False，不抛异常")

# fetch_manifest：字段合法但签名错误 → 弃用该地址并换下一个前缀（本地回环模式同样验签，无旁路）
_SIG_BAD_MJSON = json.dumps(tamper_signed(manifest(version="1.4.9", size_zip=10),
                                          signature=base64.b64encode(b"\x01" * _k).decode())).encode("utf-8")
su._last_good_prefix = ""
with mock.patch("urllib.request.urlopen",
                fake_urlopen({FP0 + su.MANIFEST_URL: lambda: FakeResp(_SIG_BAD_MJSON, ctype="application/json"),
                              FP1 + su.MANIFEST_URL: lambda: FakeResp(VALID_MJSON, ctype="application/json")})):
    m, worked = su.fetch_manifest(log=lambda s: None)
check(m["version"] == "1.4.9" and worked == FP1,
      "fetch_manifest：签名错误 → 弃用该地址并换下一个前缀（字段合法 ≠ 可信）")
# 全部前缀签名错误 → manifest，消息点名「签名校验」
all_badsig = {p + su.MANIFEST_URL: (lambda: FakeResp(_SIG_BAD_MJSON, ctype="application/json")) for p in (FP0, FP1, FP2)}
all_badsig[su.MANIFEST_URL] = lambda: FakeResp(_SIG_BAD_MJSON, ctype="application/json")
su._last_good_prefix = ""
try:
    with mock.patch("urllib.request.urlopen", fake_urlopen(all_badsig)):
        su.fetch_manifest(log=lambda s: None)
    check(False, "fetch_manifest：全部前缀签名错误应抛错")
except su.UpdateError as e:
    check(e.code == "manifest" and "签名校验" in e.message,
          f"fetch_manifest：全部签名错误 → manifest 且消息点名签名（实际 {e.code}）")
su._last_good_prefix = ""   # 还原，避免影响后续用例的前缀顺序

# --verify-manifest 入口：合法 → 0 + OK；篡改 → 1；文件不存在 → 1（子进程打桩测试公钥，不读真实私钥）
with tempfile.TemporaryDirectory() as _d:
    _good = os.path.join(_d, "latest.json")
    with open(_good, "w", encoding="utf-8") as f:
        json.dump(manifest(), f)
    _wrapper = ("import sys, importlib.util;"
                "spec = importlib.util.spec_from_file_location('setup_update', %r);"
                "su = importlib.util.module_from_spec(spec); spec.loader.exec_module(su);"
                "su.RELEASE_KEY_N = %r; su.RELEASE_KEY_E = %d;"
                "sys.exit(su._verify_manifest_cli(sys.argv[1]))"
                ) % (os.path.join(_HERE, "setup_update.py"), format(_TEST_N, "x"), _TEST_E)
    r = subprocess.run([sys.executable, "-c", _wrapper, _good], capture_output=True, text=True, encoding="utf-8", errors="replace")
    check(r.returncode == 0 and "OK" in r.stdout,
          f"--verify-manifest：合法清单 → 退出码 0 且打印 OK（实际 {r.returncode}：{r.stdout.strip()}）")
    _bad = os.path.join(_d, "bad.json")
    with open(_bad, "w", encoding="utf-8") as f:
        json.dump(tamper_signed(manifest(), sha256_zip="b" * 64), f)
    r = subprocess.run([sys.executable, "-c", _wrapper, _bad], capture_output=True, text=True, encoding="utf-8", errors="replace")
    check(r.returncode == 1 and "签名" in r.stdout,
          f"--verify-manifest：篡改清单 → 退出码 1（实际 {r.returncode}：{r.stdout.strip()}）")
    r = subprocess.run([sys.executable, "-c", _wrapper, os.path.join(_d, "ghost.json")],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    check(r.returncode == 1, f"--verify-manifest：文件不存在 → 退出码 1（实际 {r.returncode}）")

# ---------------- 2e) B3：响应头成功后中途断流（超时 / 连接重置 / IncompleteRead）→ 换下一个来源 ----------------
class CutResp(FakeResp):
    """前 cut_after 次 read 正常，之后再 read 就抛异常——模拟「200 响应头 + 首块数据后断流」"""
    def __init__(self, body, exc, cut_after=1, **kw):
        super().__init__(body, **kw)
        self._exc = exc
        self._reads = 0
        self._cut_after = cut_after
    def read(self, n=-1):
        self._reads += 1
        if self._reads > self._cut_after:
            raise self._exc
        return super().read(n)

OLD_MJSON = json.dumps(manifest(version="1.4.4", size_zip=10)).encode("utf-8")   # 测试密钥签名
NEW_MJSON = json.dumps(manifest(version="1.4.9", size_zip=10)).encode("utf-8")
# 清单：第一个镜像 200 后 TimeoutError 断流 → 第二个来源被请求并成功
su._last_good_prefix = ""
cut_calls = []
with mock.patch("urllib.request.urlopen",
                fake_urlopen({FP0 + su.MANIFEST_URL: lambda: CutResp(OLD_MJSON, TimeoutError("模拟中途断流"), ctype="application/json"),
                              FP1 + su.MANIFEST_URL: lambda: FakeResp(VALID_MJSON, ctype="application/json")},
                             cut_calls)):
    m, worked = su.fetch_manifest(log=lambda s: None)
check(m["version"] == "1.4.9" and worked == FP1 and len(cut_calls) == 2,
      f"fetch_manifest：镜像 200 后中途断流 → 换下一个来源成功（实际请求 {len(cut_calls)} 次）")
su._last_good_prefix = ""
# 清单：IncompleteRead 断流 → 同样换下一个来源
su._last_good_prefix = ""
cut_calls_ir = []
with mock.patch("urllib.request.urlopen",
                fake_urlopen({FP0 + su.MANIFEST_URL: lambda: CutResp(OLD_MJSON, http.client.IncompleteRead(b"partial"), ctype="application/json"),
                              FP1 + su.MANIFEST_URL: lambda: FakeResp(VALID_MJSON, ctype="application/json")},
                             cut_calls_ir)):
    m, worked = su.fetch_manifest(log=lambda s: None)
check(m["version"] == "1.4.9" and worked == FP1 and len(cut_calls_ir) == 2,
      f"fetch_manifest：IncompleteRead 断流 → 换下一个来源成功（实际请求 {len(cut_calls_ir)} 次）")
su._last_good_prefix = ""

# ---------------- 2f) B4：加速镜像说「没有新版本」→ 用 GitHub 直连复核一次 ----------------
# 镜像缓存旧清单（1.4.4 < 当前 1.4.5？不——「不高于当前」即说没有新版）+ 直连有新版本 → 取直连
su._last_good_prefix = ""
b4_calls = []
with mock.patch("urllib.request.urlopen",
                fake_urlopen({FP0 + su.MANIFEST_URL: lambda: FakeResp(OLD_MJSON, ctype="application/json"),
                              su.MANIFEST_URL: lambda: FakeResp(NEW_MJSON, ctype="application/json")}, b4_calls)):
    m, worked = su.fetch_manifest(current_version="1.4.5", log=lambda s: None)
check(m["version"] == "1.4.9" and worked == "",
      "fetch_manifest：镜像清单版本不高于当前 → 直连复核，以直连（有新版本）为准")
check(b4_calls == [FP0 + su.MANIFEST_URL, su.MANIFEST_URL],
      f"fetch_manifest：复核恰好多请求一次直连（实际 {b4_calls}）")
# 镜像旧 + 直连不通（国内常见）→ 沿用镜像结果，日志注明可能有缓存延迟
su._last_good_prefix = ""
b4b_calls = []
b4b_logs = []
with mock.patch("urllib.request.urlopen",
                fake_urlopen({FP0 + su.MANIFEST_URL: lambda: FakeResp(OLD_MJSON, ctype="application/json"),
                              su.MANIFEST_URL: urllib.error.URLError("direct down")}, b4b_calls)):
    m, worked = su.fetch_manifest(current_version="1.4.5", log=b4b_logs.append)
check(m["version"] == "1.4.4" and worked == FP0 and len(b4b_calls) == 2,
      "fetch_manifest：镜像无新版且直连不通 → 沿用镜像清单并保留镜像前缀")
check(any("缓存" in s for s in b4b_logs), "fetch_manifest：沿用镜像时日志写明可能有缓存延迟")
# 镜像有新版 → 立即返回，不复核直连
su._last_good_prefix = ""
b4c_calls = []
with mock.patch("urllib.request.urlopen",
                fake_urlopen({FP0 + su.MANIFEST_URL: lambda: FakeResp(NEW_MJSON, ctype="application/json")}, b4c_calls)):
    m, worked = su.fetch_manifest(current_version="1.4.5", log=lambda s: None)
check(m["version"] == "1.4.9" and worked == FP0 and len(b4c_calls) == 1,
      "fetch_manifest：镜像有新版 → 立即返回，不请求直连")
# current_version 未传 → 不复核（兼容旧调用方）
su._last_good_prefix = ""
b4d_calls = []
with mock.patch("urllib.request.urlopen",
                fake_urlopen({FP0 + su.MANIFEST_URL: lambda: FakeResp(OLD_MJSON, ctype="application/json")}, b4d_calls)):
    m, worked = su.fetch_manifest(log=lambda s: None)
check(m["version"] == "1.4.4" and worked == FP0 and len(b4d_calls) == 1,
      "fetch_manifest：未传 current_version → 不复核直连")
# 总时限已到 → 不复核
su._last_good_prefix = ""
b4e_calls = []
mono4 = iter([0.0, 0.0, 999.0])   # start=0；循环检查=0；复核剩余预算=999 → 不再请求直连
with mock.patch("urllib.request.urlopen",
                fake_urlopen({FP0 + su.MANIFEST_URL: lambda: FakeResp(OLD_MJSON, ctype="application/json")}, b4e_calls)), \
     mock.patch.object(su, "_monotonic", lambda: next(mono4)):
    m, worked = su.fetch_manifest(current_version="1.4.5", log=lambda s: None)
check(m["version"] == "1.4.4" and worked == FP0 and len(b4e_calls) == 1,
      "fetch_manifest：总时限已到 → 不再用直连复核")
su._last_good_prefix = ""   # 还原，避免影响后续用例的前缀顺序

# ---------------- 4) download_package（哈希不符换前缀；全部不符删 .part 抛 verify） ----------------
ZIP_BODY = b"PK\x03\x04" + os.urandom(4096)
ZIP_SHA = hashlib.sha256(ZIP_BODY).hexdigest()
BAD_BODY = b"PK\x03\x04" + os.urandom(1024)  # 魔数对、哈希错
P0, P1, P2 = su.MIRROR_PREFIXES
GOOD_URL_145 = GOOD_URL

def dl_manifest(url=GOOD_URL_145, body_sha=ZIP_SHA, size=None):
    m = manifest(url=url, sha256_zip=body_sha)
    m["size_zip"] = len(ZIP_BODY) if size is None else size
    return m

calls = []
responses = {
    P0 + GOOD_URL_145: FakeResp(BAD_BODY, ctype="application/octet-stream"),
    P1 + GOOD_URL_145: FakeResp(ZIP_BODY, ctype="application/octet-stream"),
}
with tempfile.TemporaryDirectory() as d:
    with mock.patch("urllib.request.urlopen", fake_urlopen(responses, calls)):
        su._last_good_prefix = ""  # 清掉别的用例可能留下的记录
        out_path = su.download_package(dl_manifest(), d, log=lambda s: None)
    check(os.path.isfile(out_path) and out_path.endswith("AttachLink-1.4.5.zip"),
          "download_package：哈希不符换下一个前缀后成功")
    check(open(out_path, "rb").read() == ZIP_BODY, "download_package：落盘内容与校验一致")
    check(not any(f.endswith(".part") for f in os.listdir(d)), "download_package：成功后无 .part 残留")
    check(calls[0] == P0 + GOOD_URL_145 and P1 + GOOD_URL_145 in calls,
          "download_package：先试第一个镜像再回退第二个")
    check(su._last_good_prefix == P1, "download_package：记录上次成功的前缀")
    # 上次成功前缀优先
    calls2 = []
    with mock.patch("urllib.request.urlopen", fake_urlopen({P1 + GOOD_URL_145: FakeResp(ZIP_BODY, ctype="application/octet-stream")}, calls2)):
        su.download_package(dl_manifest(), d, log=lambda s: None)
    check(calls2[0] == P1 + GOOD_URL_145, "download_package：上次成功的前缀优先重试")

with tempfile.TemporaryDirectory() as d:
    su._last_good_prefix = ""
    all_bad = {p + GOOD_URL_145: FakeResp(BAD_BODY, ctype="application/octet-stream") for p in (P0, P1, P2)}
    all_bad[GOOD_URL_145] = FakeResp(b"<html>proxy error</html>", ctype="text/html")  # 直连返回网页
    try:
        with mock.patch("urllib.request.urlopen", fake_urlopen(all_bad)):
            su.download_package(dl_manifest(), d, log=lambda s: None)
        check(False, "download_package：全部前缀不符应抛 UpdateError")
    except su.UpdateError as e:
        check(e.code == "verify", f"download_package：全部不符 → verify（实际 {e.code}）")
    check(os.listdir(d) == [], "download_package：全部失败后 .part 已删除、目录干净")
    # 全部是网络类失败 → network
    net_fail = {u: urllib.error.URLError("conn refused") for u in
                [p + GOOD_URL_145 for p in (P0, P1, P2)] + [GOOD_URL_145]}
    try:
        with mock.patch("urllib.request.urlopen", fake_urlopen(net_fail)):
            su.download_package(dl_manifest(), d, log=lambda s: None)
        check(False, "download_package：全部网络失败应抛 UpdateError")
    except su.UpdateError as e:
        check(e.code == "network", f"download_package：全部网络失败 → network（实际 {e.code}）")
    # 超大 Content-Length 提前拒绝（其余前缀按网络失败兜底，最终仍归为 verify）
    huge_hdr = {"Content-Type": "application/octet-stream", "Content-Length": str(su.MAX_PACKAGE_BYTES + 1)}
    huge = {P0 + GOOD_URL_145: FakeResp(ZIP_BODY, ctype="application/octet-stream", headers=huge_hdr),
            "*": urllib.error.URLError("conn refused")}
    try:
        with mock.patch("urllib.request.urlopen", fake_urlopen(huge)):
            su.download_package(dl_manifest(), d, log=lambda s: None)
        check(False, "download_package：超大 Content-Length 应拒绝")
    except su.UpdateError as e:
        check(e.code == "verify", "download_package：超大 Content-Length → verify")
    # 大小与清单不符
    wrong_size = {P0 + GOOD_URL_145: FakeResp(ZIP_BODY, ctype="application/octet-stream"),
                  "*": urllib.error.URLError("conn refused")}
    try:
        with mock.patch("urllib.request.urlopen", fake_urlopen(wrong_size)):
            su.download_package(dl_manifest(size=len(ZIP_BODY) + 1), d, log=lambda s: None)
        check(False, "download_package：大小不符应拒绝")
    except su.UpdateError as e:
        check(e.code == "verify" and "大小" in e.message, "download_package：大小不符 → verify")
    # 回环下载地址（本地测试清单）：只直连、不拼镜像前缀
    loop_url = "http://127.0.0.1:8765/AttachLink-1.4.5.zip"
    loop_calls = []
    with mock.patch("urllib.request.urlopen", fake_urlopen({loop_url: FakeResp(ZIP_BODY, ctype="application/octet-stream")}, loop_calls)):
        got = su.download_package(dl_manifest(url=loop_url), d, log=lambda s: None)
    check(os.path.isfile(got) and loop_calls == [loop_url], "download_package：回环地址只直连、成功落盘")

# B3：安装包下载中途断流（首块后连接重置）→ 换下一个来源，且无 .part 残留
with tempfile.TemporaryDirectory() as d:
    cut_dl_calls = []
    with mock.patch("urllib.request.urlopen",
                    fake_urlopen({P0 + GOOD_URL_145: lambda: CutResp(ZIP_BODY, ConnectionResetError("模拟连接重置"), ctype="application/octet-stream"),
                                  P1 + GOOD_URL_145: lambda: FakeResp(ZIP_BODY, ctype="application/octet-stream")},
                                 cut_dl_calls)):
        su._last_good_prefix = ""
        out_path = su.download_package(dl_manifest(), d, log=lambda s: None)
    check(os.path.isfile(out_path) and len(cut_dl_calls) == 2,
          f"download_package：首地址中途断流 → 换下一个来源成功（实际 {len(cut_dl_calls)} 次请求）")
    check(not any(f.endswith(".part") for f in os.listdir(d)), "download_package：断流后无 .part 残留")
su._last_good_prefix = ""

# ---------------- 5) extract_installer ----------------
EXE_BODY = b"MZ" + os.urandom(2048)
EXE_SHA = hashlib.sha256(EXE_BODY).hexdigest()
def ex_manifest(sha=EXE_SHA):
    return manifest(sha256_exe=sha)

def make_zip(path, entries):
    with zipfile.ZipFile(path, "w") as z:
        for name, data in entries:
            z.writestr(name, data)

with tempfile.TemporaryDirectory() as d:
    zp = os.path.join(d, "pkg.zip")
    make_zip(zp, [("AttachLinkSetup.exe", EXE_BODY), ("README.md", b"readme")])
    dest = os.path.join(d, "out")
    exe_path = su.extract_installer(zp, ex_manifest(), dest)
    check(os.path.isfile(exe_path) and os.path.basename(exe_path) == "AttachLinkSetup.exe",
          "extract_installer：正常取出顶层唯一 exe（basename 落盘）")
    check(open(exe_path, "rb").read() == EXE_BODY, "extract_installer：内容一致")
def assert_nothing_written(dest, label):
    """目标目录不存在或为空都算「没有写出文件」"""
    check(not os.path.exists(dest) or os.listdir(dest) == [], label)

with tempfile.TemporaryDirectory() as d:
    zp = os.path.join(d, "pkg.zip")
    make_zip(zp, [("README.md", b"readme"), ("tools/AttachLinkSetup.exe", b"not top-level")])
    dest = os.path.join(d, "out")
    expect_update_error(lambda: su.extract_installer(zp, ex_manifest(), dest), "extract",
                        "extract_installer：顶层没有 exe → extract")
    assert_nothing_written(dest, "extract_installer：失败时目标目录没有写出文件")
with tempfile.TemporaryDirectory() as d:
    zp = os.path.join(d, "pkg.zip")
    make_zip(zp, [("AttachLinkSetup.exe", b"A"), ("AttachLinkSetup.exe", b"B")])  # 同名两条目
    dest = os.path.join(d, "out")
    expect_update_error(lambda: su.extract_installer(zp, ex_manifest(), dest), "extract",
                        "extract_installer：顶层多个 exe → extract")
    assert_nothing_written(dest, "extract_installer：多 exe 时目标目录干净")
with tempfile.TemporaryDirectory() as d:
    zp = os.path.join(d, "pkg.zip")
    make_zip(zp, [("sub/../../AttachLinkSetup.exe", b"evil")])  # 条目名含 ..
    dest = os.path.join(d, "out")
    parent_before = sorted(os.listdir(d))
    expect_update_error(lambda: su.extract_installer(zp, ex_manifest(), dest), "extract",
                        "extract_installer：条目名含 .. → extract")
    check(sorted(os.listdir(d)) == parent_before,
          "extract_installer：zip-slip 条目没有写出任何目录/文件")
with tempfile.TemporaryDirectory() as d:
    zp = os.path.join(d, "pkg.zip")
    make_zip(zp, [("C:/evil/AttachLinkSetup.exe", b"evil")])  # 绝对路径条目
    dest = os.path.join(d, "out")
    expect_update_error(lambda: su.extract_installer(zp, ex_manifest(), dest), "extract",
                        "extract_installer：绝对路径条目 → extract")
    assert_nothing_written(dest, "extract_installer：绝对路径条目没有写出目标目录")
with tempfile.TemporaryDirectory() as d:
    zp = os.path.join(d, "pkg.zip")
    make_zip(zp, [("AttachLinkSetup.exe", EXE_BODY)])
    dest = os.path.join(d, "out")
    os.makedirs(dest)
    expect_update_error(lambda: su.extract_installer(zp, ex_manifest(sha="b" * 64), dest), "verify",
                        "extract_installer：exe 哈希不符 → verify")
    assert_nothing_written(dest, "extract_installer：哈希不符时已删除落盘文件")

# ---------------- 6) build_handoff_args ----------------
check(su.build_handoff_args("C:\\new\\AttachLinkSetup.exe", "C:\\old\\AttachLinkSetup.exe", 4321)
      == ["C:\\new\\AttachLinkSetup.exe", "/install-all", "--updated-from",
          "C:\\old\\AttachLinkSetup.exe", "--wait-pid", "4321"],
      "build_handoff_args：默认动作 /install-all + 交接参数")
check(su.build_handoff_args("n", "o", 1, action="/install")
      == ["n", "/install", "--updated-from", "o", "--wait-pid", "1"],
      "build_handoff_args：可指定动作")

# ---------------- 7) replace_old_exe（临时目录 + 打桩 _running_exe；返回结构化字典） ----------------
#  council C3 修复：先暂存并 SHA-256 校验内容再动旧文件；回滚失败如实报告（不再谎称「已回滚」）；
#  旧文件不存在时按 status 区分 recreated（父目录可写）/ missing（不可写），同文件 → same。
NEW_BODY = b"NEW-INSTALLER-BYTES-0123456789"
NEW_SHA = hashlib.sha256(NEW_BODY).hexdigest()
with tempfile.TemporaryDirectory() as d:
    src = os.path.join(d, "new", "AttachLinkSetup.exe")
    os.makedirs(os.path.dirname(src))
    with open(src, "wb") as f:
        f.write(NEW_BODY)

    def leftovers(dirpath):
        return [n for n in os.listdir(dirpath) if ".new-" in n or ".bak-" in n]

    with mock.patch.object(su, "_running_exe", lambda: src):
        # 旧文件不存在 + 父目录不存在 → missing，说明保留原始原因
        r = su.replace_old_exe(os.path.join(d, "ghost", "AttachLinkSetup.exe"))
        check(not r["ok"] and r["status"] == "missing" and "目录不存在" in r["message"] and r["backup"] is None,
              "replace_old_exe：旧文件不存在且父目录不存在 → missing，说明含原始原因")
        # 与自身同一路径 → same
        r = su.replace_old_exe(src)
        check(r["ok"] and r["status"] == "same" and "同一文件" in r["message"]
              and open(src, "rb").read() == NEW_BODY, "replace_old_exe：与自身同路径 → same，不替换")
        # 旧文件不存在 + 父目录可写 → recreated，新文件哈希与本安装器一致
        re_dir = os.path.join(d, "recreated")
        os.makedirs(re_dir)
        re_old = os.path.join(re_dir, "AttachLinkSetup.exe")
        r = su.replace_old_exe(re_old)
        check(r["ok"] and r["status"] == "recreated" and r["backup"] is None,
              "replace_old_exe：旧文件不存在且父目录可写 → recreated")
        check(open(re_old, "rb").read() == NEW_BODY and hashlib.sha256(open(re_old, "rb").read()).hexdigest() == NEW_SHA,
              "replace_old_exe：recreated 后新文件内容与 SHA-256 与本安装器一致")
        check(leftovers(re_dir) == [], "replace_old_exe：recreated 后无临时文件残留")
        # 旧文件不存在 + 父目录不可写（复制被拒）→ missing
        un_dir = os.path.join(d, "unwritable")
        os.makedirs(un_dir)
        def deny_copy(a, b):
            raise OSError("[Errno 13] Permission denied（模拟目录不可写）")
        with mock.patch.object(su.shutil, "copyfile", deny_copy):
            r = su.replace_old_exe(os.path.join(un_dir, "AttachLinkSetup.exe"))
        check(not r["ok"] and r["status"] == "missing" and "不可写" in r["message"] and r["backup"] is None,
              "replace_old_exe：旧文件不存在且原位置不可写 → missing，说明含不可写原因")
        check(leftovers(un_dir) == [], "replace_old_exe：missing 路径无临时文件残留")
        # 正常覆盖 → replaced
        old = os.path.join(d, "old", "AttachLinkSetup.exe")
        os.makedirs(os.path.dirname(old))
        with open(old, "wb") as f:
            f.write(b"OLD-EXE")
        r = su.replace_old_exe(old)
        check(r["ok"] and r["status"] == "replaced" and open(old, "rb").read() == NEW_BODY,
              "replace_old_exe：正常覆盖 → replaced，内容变为本安装器")
        check(r["backup"] is None and leftovers(os.path.dirname(old)) == [],
              "replace_old_exe：成功后备份已删除、无临时文件残留")
        # 复制出「等长的错误内容」→ SHA-256 校验失败：旧文件未动、无临时残留
        with open(old, "wb") as f:
            f.write(b"OLD-EXE")
        def bad_copy(a, b):
            with open(b, "wb") as fh:
                fh.write(b"X" * len(NEW_BODY))   # 等长但内容错误（旧实现只比大小会误判成功）
        with mock.patch.object(su.shutil, "copyfile", bad_copy):
            r = su.replace_old_exe(old)
        check(not r["ok"] and r["status"] == "failed" and "未被动过" in r["message"] and r["backup"] is None,
              "replace_old_exe：复制出等长错误内容 → 失败且说明旧安装器未被动过")
        check(open(old, "rb").read() == b"OLD-EXE", "replace_old_exe：内容校验失败 → 旧文件字节不变")
        check(leftovers(os.path.dirname(old)) == [], "replace_old_exe：内容校验失败 → 无临时文件残留")
        # 复制抛异常 → 同样失败、旧文件未动
        with open(old, "wb") as f:
            f.write(b"OLD-EXE")
        def boom_copy(a, b):
            raise OSError("模拟复制失败")
        with mock.patch.object(su.shutil, "copyfile", boom_copy):
            r = su.replace_old_exe(old)
        check(not r["ok"] and r["status"] == "failed" and open(old, "rb").read() == b"OLD-EXE",
              "replace_old_exe：复制异常 → 失败，旧文件未变")
        # 旧文件改名失败 → failed（可能被占用）
        real_replace = su.os.replace
        def deny_rename(a, b):
            raise OSError("模拟被占用")
        with mock.patch.object(su.os, "replace", deny_rename):
            r = su.replace_old_exe(old)
        check(not r["ok"] and r["status"] == "failed" and "改名失败" in r["message"]
              and "可能被占用" in r["message"] and r["backup"] is None,
              "replace_old_exe：旧文件改名失败 → failed（可能被占用）")
        check(open(old, "rb").read() == b"OLD-EXE", "replace_old_exe：改名失败 → 旧文件字节不变")
        check(leftovers(os.path.dirname(old)) == [], "replace_old_exe：改名失败 → 无临时文件残留")
        # 替换失败且回滚成功 → 旧文件字节不变，说明写明「已回滚，旧文件未变」
        def fail_put_replace(a, b):
            if ".new-" in a:
                raise OSError("模拟替换失败")
            return real_replace(a, b)
        with open(old, "wb") as f:
            f.write(b"OLD-EXE")
        with mock.patch.object(su.os, "replace", fail_put_replace):
            r = su.replace_old_exe(old)
        check(not r["ok"] and r["status"] == "failed" and "已回滚" in r["message"]
              and "旧文件未变" in r["message"] and r["backup"] is None,
              "replace_old_exe：替换失败且回滚成功 → failed，说明含「已回滚，旧文件未变」")
        check(open(old, "rb").read() == b"OLD-EXE" and leftovers(os.path.dirname(old)) == [],
              "replace_old_exe：回滚成功 → 旧文件字节不变、无备份残留")
        # 替换失败且回滚也失败 → 如实报告：backup 等于备份路径，说明写明原位置没有安装器
        def fail_put_and_rollback_replace(a, b):
            if ".new-" in a:
                raise OSError("模拟替换失败")
            if ".bak-" in a:
                raise OSError("模拟回滚失败")
            return real_replace(a, b)
        with open(old, "wb") as f:
            f.write(b"OLD-EXE")
        with mock.patch.object(su.os, "replace", fail_put_and_rollback_replace):
            r = su.replace_old_exe(old)
        check(not r["ok"] and r["status"] == "failed", "replace_old_exe：替换失败且回滚也失败 → failed")
        check(isinstance(r["backup"], str) and ".bak-" in r["backup"] and os.path.isfile(r["backup"]),
              "replace_old_exe：回滚也失败 → backup 为实际保留的备份路径")
        check("原位置目前没有安装器" in r["message"] and r["backup"] in r["message"] and "改回原文件名" in r["message"],
              "replace_old_exe：回滚也失败 → 说明写明原位置没有安装器且给出备份路径")
        check(not os.path.exists(old), "replace_old_exe：回滚也失败 → 原位置确实没有安装器")
        # 等待一个自己启动、很快退出的子进程
        child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(0.4)"])
        old2 = os.path.join(d, "old2", "AttachLinkSetup.exe")
        os.makedirs(os.path.dirname(old2))
        with open(old2, "wb") as f:
            f.write(b"OLD-EXE")
        r = su.replace_old_exe(old2, wait_pid=child.pid, timeout=10)
        check(r["ok"] and r["status"] == "replaced" and open(old2, "rb").read() == NEW_BODY
              and "已退出" in r["message"], "replace_old_exe：wait_pid 等子进程退出后替换")
        child.wait()
        # 等待超时 → 不替换
        sleeper = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(3)"])
        old3 = os.path.join(d, "old3", "AttachLinkSetup.exe")
        os.makedirs(os.path.dirname(old3))
        with open(old3, "wb") as f:
            f.write(b"OLD-EXE")
        r = su.replace_old_exe(old3, wait_pid=sleeper.pid, timeout=1)
        check(not r["ok"] and r["status"] == "failed" and "超时" in r["message"]
              and open(old3, "rb").read() == b"OLD-EXE",
              "replace_old_exe：等待超时 → failed，不替换并说明原因")
        sleeper.kill()  # 只结束本测试自己启动并记录了 PID 的进程
        sleeper.wait()

# ---------------- 7b) NB-4：更新临时目录按次隔离（GUID 子目录 + 只清 24 小时旧子目录） ----------------
import time  # noqa: E402
with tempfile.TemporaryDirectory() as d:
    # 打桩 gettempdir 让根目录落在临时目录里（不碰真实 %TEMP%）
    with mock.patch.object(su.tempfile, "gettempdir", lambda: d):
        a = su._update_tmp_dir()
        b = su._update_tmp_dir()
        check(a != b and os.path.dirname(a) == os.path.dirname(b) and os.path.isdir(a) and os.path.isdir(b),
              "_update_tmp_dir：两次调用拿到不同 GUID 子目录（并发更新互不干扰）")
        # 伪造一个过期子目录（25 小时前）与一个新鲜子目录
        stale = os.path.join(os.path.dirname(a), "stale-old")
        fresh = os.path.join(os.path.dirname(a), "fresh-new")
        os.makedirs(stale); os.makedirs(fresh)
        old_time = time.time() - 25 * 3600
        os.utime(stale, (old_time, old_time))
        removed = su._clean_stale_tmp_dirs(os.path.dirname(a))
        check(removed == 1 and not os.path.exists(stale) and os.path.exists(fresh)
              and os.path.exists(a) and os.path.exists(b),
              "_clean_stale_tmp_dirs：只清超过 24 小时的旧子目录，新子目录不被清理")

# ---------------- 8) run_update（打桩网络与版本，验证汇总文本） ----------------
import setup_actions as sa_mod  # 已装 setup_update 的兄弟模块（sys.path 含 packaging）

# 打成真 zip 的发布包字节（extract_installer 会真的解包）
_buf = io.BytesIO()
with zipfile.ZipFile(_buf, "w") as _z:
    _z.writestr("AttachLinkSetup.exe", EXE_BODY)
    _z.writestr("README.md", b"readme")
PKG_BODY = _buf.getvalue()
PKG_SHA = hashlib.sha256(PKG_BODY).hexdigest()

def serve_update(responses):
    return mock.patch("urllib.request.urlopen", fake_urlopen(responses))

with tempfile.TemporaryDirectory() as d:
    # 响应按需新建（可调用对象）：run_update 会被多个用例反复调用，单实例 FakeResp 的缓冲区只够读一次
    responses_by_kind = {
        "manifest": lambda: FakeResp(json.dumps(manifest(version="1.4.5", url=GOOD_URL,
                                                         sha256_zip=PKG_SHA, sha256_exe=EXE_SHA,
                                                         size_zip=len(PKG_BODY))).encode()),
        "pkg": lambda: FakeResp(PKG_BODY, ctype="application/octet-stream"),
    }

    def route(request, timeout=None):
        url = request.full_url
        item = responses_by_kind["manifest"] if url.endswith("latest.json") else responses_by_kind["pkg"]
        return item()

    with mock.patch.object(sa_mod, "read_version", lambda: "1.4.4"), \
         mock.patch.object(su, "_update_tmp_dir", lambda: os.path.join(d, "upd")), \
         mock.patch("urllib.request.urlopen", route):
        summary = su.run_update(dry_run=True, out=lambda s: None)
    upd_dir = os.path.join(d, "upd")
    exe_in_tmp = os.path.join(upd_dir, "AttachLinkSetup.exe")
    check(summary.startswith("试运行完成") and "v1.4.5" in summary and exe_in_tmp in summary,
          "run_update：dry_run → 试运行完成文本带版本与安装器路径")
    check(os.path.isfile(exe_in_tmp) and open(exe_in_tmp, "rb").read() == EXE_BODY,
          "run_update：dry_run 解压出的安装器内容正确")
    head16 = summary.split("SHA-256 前16位 ")[1].split("）")[0]
    check(head16 == hashlib.sha256(EXE_BODY).hexdigest()[:16], "run_update：试运行文本含 SHA-256 前 16 位")
    # 已是最新
    with mock.patch.object(sa_mod, "read_version", lambda: "1.4.5"), \
         mock.patch.object(su, "_update_tmp_dir", lambda: os.path.join(d, "upd")), \
         mock.patch("urllib.request.urlopen", route):
        summary = su.run_update(dry_run=True, out=lambda s: None)
    check(summary == "当前已是最新版本 v1.4.5", "run_update：版本相同 → 已是最新文本")
    # 检查失败
    def always_fail(request, timeout=None):
        raise urllib.error.URLError("no route")
    with mock.patch.object(sa_mod, "read_version", lambda: "1.4.4"), \
         mock.patch.object(su, "_update_tmp_dir", lambda: os.path.join(d, "upd2")), \
         mock.patch("urllib.request.urlopen", always_fail):
        summary = su.run_update(dry_run=True, out=lambda s: None)
    check(summary.startswith("失败：") and "network" in summary and "github.com/jasonisbrave/attachlink/releases" in summary,
          "run_update：网络失败 → 「失败：」文本带类别与建议")
    check(not os.path.exists(os.path.join(d, "upd2")), "run_update：失败后临时目录已清理")
    # 交接仅支持发布版：源码运行（无 sys.frozen）时拒绝交接
    with mock.patch.object(sa_mod, "read_version", lambda: "1.4.4"), \
         mock.patch.object(su, "_update_tmp_dir", lambda: os.path.join(d, "upd3")), \
         mock.patch("urllib.request.urlopen", route):
        summary = su.run_update(dry_run=False, out=lambda s: None)
    check(summary.startswith("失败：") and "发布版安装器" in summary,
          "run_update：非冻结环境拒绝交接（防止误把 python.exe 复制到旧路径）")
    # B5：交接启动成功后删除本次下载的 zip（解压出的新安装器正在运行，必须保留）
    with mock.patch.object(sa_mod, "read_version", lambda: "1.4.4"), \
         mock.patch.object(su, "_update_tmp_dir", lambda: os.path.join(d, "upd4")), \
         mock.patch.object(su, "launch_handoff", lambda args: 4321), \
         mock.patch.object(su.sys, "frozen", True, create=True), \
         mock.patch("urllib.request.urlopen", route):
        summary4 = su.run_update(dry_run=False, out=lambda s: None)
    check(summary4.startswith("已启动新版本安装器（PID 4321）"), "run_update：交接启动成功文本（B5 前置）")
    upd4 = os.path.join(d, "upd4")
    check(not os.path.exists(os.path.join(upd4, "AttachLink-1.4.5.zip")),
          "run_update：交接成功后本次下载的 zip 已删除（B5）")
    check(os.path.isfile(os.path.join(upd4, "AttachLinkSetup.exe")),
          "run_update：交接成功后解压出的新安装器保留（正在运行）")

print("RESULT:", "PASS" if ok else "FAIL")

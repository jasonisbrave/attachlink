// =====================================================================
// AttachLink.Watchdog：判定逻辑（纯函数，不碰 COM）
// ---------------------------------------------------------------------
// 移植自 scripts\watchdog.ps1 的 Get-PresenceDecision / Get-HiddenDecision / Test-HiddenReevaluate /
// Get-BackoffDelay / Get-InstanceDecision / Test-SamePath。由 --selftest 逐条回归（用例照搬
// scripts\test-watchdog-units.ps1 的 G / H 段与在位判定假对象用例）。
// =====================================================================
using System;
using System.IO;

namespace AttachLink.Watchdog
{
    internal enum LookupState { Found, NotFound, Error }

    /// <summary>在位判定的五种结论。</summary>
    internal enum Presence
    {
        /// <summary>菜单不在位 → 走装载流程。</summary>
        NotLoaded,
        /// <summary>菜单在位且在位的就是配置的 xlam → 标记完成。</summary>
        InPlace,
        /// <summary>菜单在位但在位的是异路径同名副本 → 不动它。</summary>
        Foreign,
        /// <summary>菜单在位但加载项工作簿明确不在位（命令栏缓存残留的「死菜单」）→ 按未装载处理。</summary>
        DeadMenu,
        /// <summary>菜单在位但查询出错 / 路径读不到 → 不标 seen，下周期重试。</summary>
        Unknown
    }

    internal enum HiddenVerdict { Process, Wait, Skip }

    /// <summary>某个文档窗口在「隐藏判定」上的状态（按窗口键保存，窗口消失即丢弃）。</summary>
    internal sealed class WindowState
    {
        public DateTime? Since;
        public bool Done;
        public int Rev;
        public bool Shown;
        public DateTime? LastCheck;
        public bool Logged;
        public DateTime? UnknownSince;
        public bool UnknownEscalated;
    }

    internal static class Decisions
    {
        public const string AddInName = "AttachLink.xlam";
        public const string MenuTag = "AttachLinkCtx";

        /// <summary>菜单 Tag 是否在位 + Workbooks 三态查询 → 五种结论。</summary>
        public static Presence GetPresenceDecision(bool menuExists, LookupState lookupState, bool samePath)
        {
            if (!menuExists) return Presence.NotLoaded;
            if (lookupState == LookupState.Found) return samePath ? Presence.InPlace : Presence.Foreign;
            if (lookupState == LookupState.NotFound) return Presence.DeadMenu;
            return Presence.Unknown;
        }

        /// <summary>附着失败后的下次重试间隔（秒）：首次取 base，之后逐次翻倍，封顶 max。prev ≤ 0 视为首次；base 大于 max 时以 max 为准。</summary>
        public static double GetBackoffDelay(double prevSeconds, double baseSeconds, double maxSeconds)
        {
            double d = prevSeconds > 0 ? prevSeconds * 2 : baseSeconds;
            if (d < baseSeconds) d = baseSeconds;
            if (d > maxSeconds) d = maxSeconds;
            return d;
        }

        /// <summary>实例可见性判定（v1.6.2：只附着用户看得见的会话）。appVisible：1 = Application.Visible 为真；0 = 为假；-1 = 读不到。
        /// 返回 true = attach（按工作窗口处理）；false = defer（本周期不处理、不标 seen，下周期再看）。
        /// allowHidden（仅回归用）= 旧规则：不可见实例首次出现只记时间，宽限一个周期后照常处理；读不到按可见处理。</summary>
        public static bool GetInstanceDecision(int appVisible, bool allowHidden, DateTime? firstSeen, DateTime now, int intervalSeconds)
        {
            if (appVisible == 1) return true;
            if (allowHidden)
            {
                if (appVisible < 0) return true;
                if (!firstSeen.HasValue) return false;
                if ((now - firstSeen.Value).TotalSeconds < intervalSeconds) return false;
                return true;
            }
            return false;
        }

        /// <summary>对一个尚未标记 seen 的窗口，按 Excel 报告的 Window.Visible（1 / 0，-1 = 读不到）决定本周期怎么做。
        /// st 由调用方按窗口键保存；读不到按「未知」处理，只延迟、重试。</summary>
        public static HiddenVerdict GetHiddenDecision(WindowState st, int winVisible, bool shownNow, DateTime now, int intervalSeconds)
        {
            if (winVisible < 0) return HiddenVerdict.Wait;
            st.UnknownSince = null;
            st.UnknownEscalated = false;
            if (winVisible != 0) return HiddenVerdict.Process;
            // Excel 报告为隐藏：宽限一个周期再判定（用户的工作窗口在创建瞬间可能尚未显示）
            if (!st.Since.HasValue) { st.Since = now; return HiddenVerdict.Wait; }
            if ((now - st.Since.Value).TotalSeconds < intervalSeconds) return HiddenVerdict.Wait;
            st.Done = true;
            st.Shown = shownNow;
            st.LastCheck = now;
            return HiddenVerdict.Skip;
        }

        /// <summary>对一个已标记 seen、已判定为隐藏的窗口，每个周期只用 Win32 可见性决定是否重新评估。
        /// 返回 true 时已把状态复位为待评估。任何路径都不会让一个窗口被永久跳过。</summary>
        public static bool TestHiddenReevaluate(WindowState st, bool shownNow, DateTime now, int maxReeval, int recheckSeconds)
        {
            bool turned = shownNow && !st.Shown;
            st.Shown = shownNow;
            if (!shownNow) return false;
            double sinceCheck = double.MaxValue;
            if (st.LastCheck.HasValue) sinceCheck = (now - st.LastCheck.Value).TotalSeconds;
            if (!((turned && st.Rev < maxReeval) || sinceCheck >= recheckSeconds)) return false;
            if (turned && st.Rev < maxReeval) st.Rev++;   // 计数不越过上限
            st.LastCheck = now;
            st.Done = false;
            st.Since = now;
            return true;
        }

        /// <summary>两个路径是否指向同一文件（FullPath 后不区分大小写比较）；任一侧无效返回 false。</summary>
        public static bool SamePath(string a, string b)
        {
            try { return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase); }
            catch { return false; }
        }
    }
}

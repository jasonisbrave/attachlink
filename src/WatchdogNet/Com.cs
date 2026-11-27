// =====================================================================
// AttachLink.Watchdog（预研基础第 5 步）：Win32 / oleacc 辅助与「晚绑定 COM + 确定性释放」
// ---------------------------------------------------------------------
// 这是 scripts\watchdog.ps1 里 C# 辅助块（AttachLinkWatchdog.Native）与 COM 纪律（Release-Com、
// Menu-Exists、Get-LoadedAddInPath 里的下标遍历与逐个释放）的移植。行为必须与 PowerShell 版逐项等价。
// COM 纪律：看门狗碰过的每个 COM 对象都由 ComScope 记下并在用完时显式释放（循环 ReleaseComObject 到归零）；
// 宿主在 Quit 之后要等外部 COM 引用全部释放才真正退出，不能依赖 GC（WORKLOG #27）。
//
// 晚绑定直接走 IDispatch::GetIDsOfNames / Invoke（与脚本宿主同一个通道），不用 Type.InvokeMember：
//   2026-11-27 实测，Type.InvokeMember 读 Workbook.FullName 时传 CultureInfo.InvariantCulture 会得到
//   TYPE_E_INVDATAREAD（0x80028018「格式太旧或是类型库无效」），传 CurrentCulture 才行；而 Application、
//   CommandBars 上的同一写法却正常。原因在 CLR 这一层会去碰类型信息。直接调 IDispatch 不碰类型库，也没有这个坑。
// =====================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace AttachLink.Watchdog
{
    internal static class Native
    {
        private delegate bool EnumWindowsProc(IntPtr h, IntPtr lp);

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lp);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc cb, IntPtr lp);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr h);
        [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr h, uint dwId, ref Guid riid,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);

        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr CreateToolhelp32Snapshot(uint flags, uint pid);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern bool Process32FirstW(IntPtr h, ref ProcessEntry32 lp);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern bool Process32NextW(IntPtr h, ref ProcessEntry32 lp);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr h);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct ProcessEntry32
        {
            public uint dwSize; public uint cntUsage; public uint th32ProcessID; public IntPtr th32DefaultHeapID;
            public uint th32ModuleID; public uint cntThreads; public uint th32ParentProcessID; public int pcPriClassBase; public uint dwFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szExeFile;
        }

        private static readonly Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");
        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;

        private static bool IsClass(IntPtr h, string cls)
        {
            StringBuilder sb = new StringBuilder(64);
            GetClassName(h, sb, 64);
            return sb.ToString() == cls;
        }

        /// <summary>EXCEL7 窗口所属进程的 PID。</summary>
        public static uint GetWindowProcessId(long hwnd)
        {
            uint pid;
            GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
            return pid;
        }

        /// <summary>枚举 XLMAIN 顶层窗口下的 EXCEL7 文档窗口句柄（含隐藏的；是否处理由主循环按 Window.Visible 决定）。</summary>
        public static List<long> FindExcel7Windows()
        {
            List<long> list = new List<long>();
            EnumWindows(delegate(IntPtr top, IntPtr lp)
            {
                if (IsClass(top, "XLMAIN"))
                {
                    EnumChildWindows(top, delegate(IntPtr child, IntPtr lp2)
                    {
                        if (IsClass(child, "EXCEL7")) list.Add(child.ToInt64());
                        return true;
                    }, IntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);
            return list;
        }

        /// <summary>目标窗口是否仍然存在（写操作前的复核用）。</summary>
        public static bool IsWindowAlive(long hwnd) { return IsWindow(new IntPtr(hwnd)); }

        /// <summary>窗口自身是否带 WS_VISIBLE（纯 Win32，不涉及宿主）。实例整体不可见时所有文档窗口在 Win32 层面都不可见。</summary>
        public static bool IsWindowShown(long hwnd) { return IsWindowVisible(new IntPtr(hwnd)); }

        /// <summary>从 EXCEL7 窗口取 Excel Window 对象（IDispatch）。取不到返回 null。调用方负责释放。</summary>
        public static object GetExcelWindow(long hwnd)
        {
            object win;
            Guid iid = IID_IDispatch;
            int hr = AccessibleObjectFromWindow(new IntPtr(hwnd), OBJID_NATIVEOM, ref iid, out win);
            if (hr != 0 || win == null) return null;
            return win;
        }

        /// <summary>父进程的映像名（如 sjp7.exe）；查不到返回 null。仅供父进程白名单使用（Toolhelp32，不依赖 WMI）。</summary>
        public static string GetParentProcessName(uint pid)
        {
            IntPtr snap = CreateToolhelp32Snapshot(2 /*TH32CS_SNAPPROCESS*/, 0);
            if (snap == new IntPtr(-1)) return null;
            try
            {
                ProcessEntry32 e = new ProcessEntry32();
                e.dwSize = (uint)Marshal.SizeOf(typeof(ProcessEntry32));
                uint parent = 0; bool found = false;
                if (Process32FirstW(snap, ref e))
                {
                    do
                    {
                        if (e.th32ProcessID == pid) { parent = e.th32ParentProcessID; found = true; break; }
                    } while (Process32NextW(snap, ref e));
                }
                if (!found || parent == 0) return null;
                e = new ProcessEntry32();
                e.dwSize = (uint)Marshal.SizeOf(typeof(ProcessEntry32));
                if (Process32FirstW(snap, ref e))
                {
                    do
                    {
                        if (e.th32ProcessID == parent) return e.szExeFile;
                    } while (Process32NextW(snap, ref e));
                }
                return null;
            }
            finally { CloseHandle(snap); }
        }
    }

    /// <summary>宿主调用失败。Message 是「操作名: 根因消息」（根因消息与 PowerShell 版日志里「附着失败 … : …」后面那段相同），
    /// InnerHResult 是根因的 HRESULT（如 DISP_E_BADINDEX 0x8002000B =「按名取不到工作簿」）。</summary>
    internal sealed class HostException : Exception
    {
        public HostException(string operation, Exception inner)
            : base(operation + ": " + (inner == null ? "?" : inner.Message))
        {
            InnerHResult = inner == null ? 0 : inner.HResult;
        }
        public HostException(string message) : base(message) { InnerHResult = 0; }
        public int InnerHResult { get; private set; }
    }

    /// <summary>一次附着里碰过的全部 COM 对象：Track 记下，Release / Dispose 显式释放（后进先出）。同一个对象只释放一次。</summary>
    internal sealed class ComScope : IDisposable
    {
        private readonly List<object> _owned = new List<object>();

        public T Track<T>(T obj)
        {
            if (obj != null && Marshal.IsComObject(obj) && !_owned.Contains(obj)) _owned.Add(obj);
            return obj;
        }

        /// <summary>立刻释放某一个对象（例如循环里的单个命令栏控件），不再等到 Dispose。</summary>
        public void Release(object obj)
        {
            if (obj == null || !Marshal.IsComObject(obj)) return;
            int i = _owned.IndexOf(obj);
            if (i >= 0) _owned.RemoveAt(i);
            ReleaseAll(obj);
        }

        private static void ReleaseAll(object o)
        {
            try { while (Marshal.ReleaseComObject(o) > 0) { } } catch { }
        }

        public void Dispose()
        {
            for (int i = _owned.Count - 1; i >= 0; i--) ReleaseAll(_owned[i]);
            _owned.Clear();
        }
    }

    /// <summary>对 IDispatch 对象的晚绑定访问：直接调 IDispatch::GetIDsOfNames / Invoke（见文件头的说明）。</summary>
    internal static class Com
    {
        private const ushort DispatchMethod = 1, DispatchPropertyGet = 2, DispatchPropertyPut = 4;
        private const int DISP_E_EXCEPTION = unchecked((int)0x80020009);
        private const int VariantSize = 16;            // VARIANT 的大小（x86 / x64 都是 16 字节）
        private static readonly Guid IIDNull = Guid.Empty;

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPPARAMS { public IntPtr rgvarg; public IntPtr rgdispidNamedArgs; public uint cArgs; public uint cNamedArgs; }

        [StructLayout(LayoutKind.Sequential)]
        private struct EXCEPINFO
        {
            public ushort wCode; public ushort wReserved; public IntPtr bstrSource; public IntPtr bstrDescription;
            public IntPtr bstrHelpFile; public uint dwHelpContext; public IntPtr pvReserved; public IntPtr pfnDeferredFillIn; public int scode;
        }

        [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDispatch
        {
            // 四个方法必须按 IDispatch 的 vtable 顺序全部声明（哪怕不用前两个），否则调用会落到错误的槽位上
            [PreserveSig] int GetTypeInfoCount(out uint pctinfo);
            [PreserveSig] int GetTypeInfo(uint iTInfo, int lcid, out IntPtr ppTInfo);
            [PreserveSig] int GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, uint cNames, int lcid,
                [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
            [PreserveSig] int Invoke(int dispIdMember, ref Guid riid, int lcid, ushort wFlags, ref DISPPARAMS pDispParams,
                IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr);
        }

        [DllImport("oleaut32.dll")] private static extern int VariantClear(IntPtr pvarg);
        [DllImport("oleaut32.dll")] private static extern int SysFreeString(IntPtr bstr);

        /// <summary>取属性（无参数）。</summary>
        public static object Get(object obj, string name) { return Invoke(obj, name, DispatchPropertyGet, null); }

        /// <summary>带参数的属性或方法（集合的 Item 要用「属性读 | 方法」才会命中，只用方法标志是 DISP_E_MEMBERNOTFOUND——2026-11-27 实测）。</summary>
        public static object GetOrInvoke(object obj, string name, params object[] args) { return Invoke(obj, name, (ushort)(DispatchPropertyGet | DispatchMethod), args); }

        /// <summary>调用方法。</summary>
        public static object Call(object obj, string name, params object[] args) { return Invoke(obj, name, DispatchMethod, args); }

        /// <summary>设置属性（只用于标量值）。</summary>
        public static void Set(object obj, string name, object value) { Invoke(obj, name, DispatchPropertyPut, new object[] { value }); }

        private static object Invoke(object obj, string name, ushort flags, object[] args)
        {
            if (obj == null) throw new HostException(name, null);
            IDispatch d = (IDispatch)obj;
            Guid iid = IIDNull;
            int[] ids = new int[1];
            int hr = d.GetIDsOfNames(ref iid, new string[] { name }, 1, 0 /*LOCALE_SYSTEM_DEFAULT*/, ids);
            if (hr != 0) throw new HostException(name, Marshal.GetExceptionForHR(hr));

            DISPPARAMS dp = new DISPPARAMS();
            IntPtr block = IntPtr.Zero;
            int argc = args == null ? 0 : args.Length;
            if (argc > 0)
            {
                // 每个参数一个 VARIANT，按从右到左的顺序存放
                block = Marshal.AllocCoTaskMem(VariantSize * argc);
                for (int i = 0; i < argc; i++) Marshal.GetNativeVariantForObject(args[argc - 1 - i] ?? System.Reflection.Missing.Value, block + VariantSize * i);
                dp.rgvarg = block; dp.cArgs = (uint)argc;
            }
            int excepSize = Marshal.SizeOf(typeof(EXCEPINFO));
            IntPtr result = Marshal.AllocCoTaskMem(VariantSize);
            IntPtr pExcep = Marshal.AllocCoTaskMem(excepSize);
            try
            {
                for (int i = 0; i < VariantSize; i++) Marshal.WriteByte(result, i, 0);
                for (int i = 0; i < excepSize; i++) Marshal.WriteByte(pExcep, i, 0);
                int argErr;
                hr = d.Invoke(ids[0], ref iid, 0, flags, ref dp, result, pExcep, out argErr);
                if (hr != 0)
                {
                    // Excel 的「找不到成员 / 下标越界」等以 DISP_E_EXCEPTION 返回，真正的 HRESULT 在 EXCEPINFO.scode 里；
                    // Type.InvokeMember 与脚本宿主都会把它折成带那个 HRESULT 的异常——这里照做
                    int scode = hr; string description = null;
                    if (hr == DISP_E_EXCEPTION)
                    {
                        EXCEPINFO ei = (EXCEPINFO)Marshal.PtrToStructure(pExcep, typeof(EXCEPINFO));
                        if (ei.scode != 0) scode = ei.scode; else if (ei.wCode != 0) scode = ei.wCode;
                        if (ei.bstrDescription != IntPtr.Zero) { description = Marshal.PtrToStringBSTR(ei.bstrDescription); SysFreeString(ei.bstrDescription); }
                        if (ei.bstrSource != IntPtr.Zero) SysFreeString(ei.bstrSource);
                        if (ei.bstrHelpFile != IntPtr.Zero) SysFreeString(ei.bstrHelpFile);
                    }
                    Exception inner = string.IsNullOrEmpty(description) ? Marshal.GetExceptionForHR(scode) : new COMException(description, scode);
                    throw new HostException(name, inner);
                }
                for (int i = 0; i < VariantSize; i++) { if (Marshal.ReadByte(result, i) != 0) { object o = Marshal.GetObjectForNativeVariant(result); VariantClear(result); return o; } }
                return null;            // VT_EMPTY：过程调用没有返回值
            }
            finally
            {
                Marshal.FreeCoTaskMem(result); Marshal.FreeCoTaskMem(pExcep);
                if (block != IntPtr.Zero) Marshal.FreeCoTaskMem(block);
            }
        }
    }
}

// =====================================================================
// AttachLink COM 引导器（Bootstrap Shim）v1.2
// ---------------------------------------------------------------------
// 背景：鼎信诺等审计软件通过 COM 自动化（CreateObject）拉起 Excel 时，
// 注册的 .xlam 加载项与 XLSTART 文件都不会被自动加载（本机已实证）。
// 本引导器是 COM 加载项（所有启动方式下都会被尝试装载），职责单一：
// 在宿主就绪后把 AttachLink.xlam 程序化装载进当前实例，
// 之后全部逻辑仍在 VBA（AttachLink.bas）中，由其 Workbook_Open 完成初始化。
//
// 兼容性设计（经实测迭代）：
//   - 同时实现两个 IDTExtensibility2 接口 IID：
//       DC1ACE6A-A7C0-49AF-9C06-69E8BD7C4F35（Extensibility PIA 官方，dual）
//       B65AD801-ABAF-11D0-BB8B-00A0C90F2744（MSADDNDR 别名 dispinterface）
//     无论宿主 QI 哪个 IID 都能命中（Excel 启动与手动连接的行为不同源）
//   - 装载动作：Startup 连接 → OnStartupComplete；AfterStartup 连接 → 直接装
//   - 零工作簿实例中 AddIns.Add 报 0x800A03EC（实测），先建临时工作簿再装载
//   - 所有回调吞掉异常并写日志（按 PID 分文件）
//
// 编译：系统自带 csc.exe（.NET Framework 4），无需安装 SDK；
// 注册：install.ps1 用 RegAsm 生成规范脚本重写为 HKCU 导入（无需管理员）。
// 注意：若宿主 Excel 以管理员（高完整性）运行，Windows 会忽略 HKCU 的
// COM 类注册，此时需要 HKLM 注册（见 README「鼎信诺共存」章节）。
// =====================================================================
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyVersion("1.0.2.0")]

namespace AttachLinkShim
{
    // Extensibility PIA 官方 IID（dual）
    [ComImport]
    [Guid("DC1ACE6A-A7C0-49AF-9C06-69E8BD7C4F35")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            int ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            ref Array custom);

        [DispId(2)]
        void OnDisconnection(int RemoveMode, ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(ref Array custom);

        [DispId(4)]
        void OnStartupComplete(ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(ref Array custom);
    }

    // MSADDNDR 别名对应的 dispinterface IID
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2Dispatch
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            int ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            ref Array custom);

        [DispId(2)]
        void OnDisconnection(int RemoveMode, ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(ref Array custom);

        [DispId(4)]
        void OnStartupComplete(ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(ref Array custom);
    }

    [ComVisible(true)]
    [Guid("D4F8A1C2-93B7-4E5A-8C6D-2F1E9B7A4C58")]
    [ProgId("AttachLinkShim.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDTExtensibility2))]
    public class Connect : IDTExtensibility2, IDTExtensibility2Dispatch
    {
        private object _app;
        private int _connectMode = -1;
        private bool _loadAttempted;

        private static void Log(string msg)
        {
            try
            {
                File.AppendAllText(
                    Path.Combine(Path.GetTempPath(), "AttachLinkShim_" + Process.GetCurrentProcess().Id + ".log"),
                    DateTime.Now.ToString("HH:mm:ss.fff ") + msg + Environment.NewLine);
            }
            catch { }
        }

        static Connect()
        {
            Log("cctor: 程序集已加载, pid=" + Process.GetCurrentProcess().Id);
        }

        public Connect()
        {
            Log("instance ctor");
        }

        private void HandleOnConnection(object Application, int ConnectMode)
        {
            Log("OnConnection enter, mode=" + ConnectMode);
            try
            {
                _app = Application;
                _connectMode = ConnectMode;
                // ext_cm_AfterStartup = 0：此时不会再收到 OnStartupComplete，直接装载
                if (ConnectMode == 0) LoadXlam();
            }
            catch (Exception ex)
            {
                Log("OnConnection err: " + ex);
            }
        }

        private void HandleOnStartupComplete()
        {
            Log("OnStartupComplete enter");
            try { LoadXlam(); }
            catch (Exception ex) { Log("OnStartupComplete err: " + ex); }
        }

        private void LoadXlam()
        {
            if (_loadAttempted) return;
            _loadAttempted = true;
            if (_app == null) { Log("LoadXlam: app 为空，跳过"); return; }
            string xlam = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Microsoft", "AddIns", "AttachLink.xlam");
            Log("LoadXlam: path=" + xlam + ", exists=" + File.Exists(xlam));
            if (!File.Exists(xlam)) return;

            dynamic app = (dynamic)_app;
            dynamic wbs = app.Workbooks;
            dynamic tempWb = null;
            if (wbs.Count == 0)
            {
                // 零工作簿实例中 AddIns.Add 会报 0x800A03EC（实测证实）：
                // 先创建临时工作簿，装载完成后再关闭（SaveChanges=false，无提示）
                Log("LoadXlam: 实例无工作簿，创建临时工作簿");
                tempWb = wbs.Add();
            }
            try
            {
                dynamic addins = app.AddIns;
                dynamic ai = addins.Add(xlam, false);
                if (!(bool)ai.Installed) ai.Installed = true;
                Log("LoadXlam: 完成");
            }
            finally
            {
                if (tempWb != null)
                {
                    try { tempWb.Close(false); Log("LoadXlam: 临时工作簿已关闭"); }
                    catch (Exception ex) { Log("LoadXlam: 临时工作簿关闭失败 " + ex.Message); }
                }
            }
        }

        void IDTExtensibility2.OnConnection(object Application, int ConnectMode, object AddInInst, ref Array custom)
        {
            HandleOnConnection(Application, ConnectMode);
        }

        void IDTExtensibility2.OnDisconnection(int RemoveMode, ref Array custom)
        {
            Log("OnDisconnection, mode=" + RemoveMode);
        }

        void IDTExtensibility2.OnAddInsUpdate(ref Array custom) { }

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            HandleOnStartupComplete();
        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
            Log("OnBeginShutdown");
        }

        void IDTExtensibility2Dispatch.OnConnection(object Application, int ConnectMode, object AddInInst, ref Array custom)
        {
            HandleOnConnection(Application, ConnectMode);
        }

        void IDTExtensibility2Dispatch.OnDisconnection(int RemoveMode, ref Array custom)
        {
            Log("OnDisconnection(dispatch), mode=" + RemoveMode);
        }

        void IDTExtensibility2Dispatch.OnAddInsUpdate(ref Array custom) { }

        void IDTExtensibility2Dispatch.OnStartupComplete(ref Array custom)
        {
            HandleOnStartupComplete();
        }

        void IDTExtensibility2Dispatch.OnBeginShutdown(ref Array custom)
        {
            Log("OnBeginShutdown(dispatch)");
        }
    }
}

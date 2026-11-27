// =====================================================================
// AttachLink.Watchdog：日志
// ---------------------------------------------------------------------
// 与 PowerShell 版同一个文件、同一套措辞、同一种行格式（「MM-dd HH:mm:ss 内容」，UTF-8 无 BOM，CRLF）——
// healthcheck.ps1 依赖其中的关键字与心跳判定。
// 新增（第 5 步确认）：大小上限，默认 2 MB；超过时把当前文件转成「.prev.log」（只保留一份）再写。
// 启动器也会往这个文件写「启动器:」行（打开-写-关闭），轮转时若文件正被占用就等下一次再转，不丢行。
// =====================================================================
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace AttachLink.Watchdog
{
    internal sealed class Logger
    {
        public const long DefaultMaxBytes = 2L * 1024 * 1024;
        private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);
        private readonly object _gate = new object();
        private readonly string _path;
        private readonly long _maxBytes;
        private readonly bool _echo;

        public Logger(string path, long maxBytes, bool echo)
        {
            _path = path;
            _maxBytes = maxBytes <= 0 ? long.MaxValue : maxBytes;
            _echo = echo;
        }

        public string Path { get { return _path; } }
        public string PreviousPath { get { return PreviousFileName(_path); } }

        /// <summary>日志主文件对应的旧文件名：AttachLinkWatchdog.log → AttachLinkWatchdog.prev.log（其余在扩展名前加 .prev）。</summary>
        public static string PreviousFileName(string path)
        {
            string dir = System.IO.Path.GetDirectoryName(path);
            string file = System.IO.Path.GetFileName(path);
            string ext = System.IO.Path.GetExtension(file);
            string stem = ext.Length > 0 ? file.Substring(0, file.Length - ext.Length) : file;
            return System.IO.Path.Combine(dir, stem + ".prev" + ext);
        }

        public void Log(string message)
        {
            string line = DateTime.Now.ToString("MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + message;
            lock (_gate)
            {
                Write(line);
            }
            if (_echo)
            {
                try { Console.WriteLine(line); } catch { }
            }
        }

        private void Write(string line)
        {
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    RotateIfNeeded();
                    File.AppendAllText(_path, line + "\r\n", Utf8NoBom);
                    return;
                }
                catch (IOException) { Thread.Sleep(50); }              // 另一个写入者（启动器）正占着文件：稍后重试
                catch (UnauthorizedAccessException) { return; }         // 目录不可写：与 PowerShell 版一致，静默放弃这一行
            }
        }

        private void RotateIfNeeded()
        {
            FileInfo fi = new FileInfo(_path);
            if (!fi.Exists || fi.Length <= _maxBytes) return;
            string prev = PreviousPath;
            if (File.Exists(prev)) File.Delete(prev);
            File.Move(_path, prev);                                     // 文件被别的进程占用时抛 IOException：下一次再转
        }
    }
}

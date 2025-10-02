//[10/01/2025]:Raksha- Minimal standalone logger for InventorAddIn
using System;
using System.IO;
using System.Text;

namespace PanelSync.InventorAddIn
{
    internal interface ILog
    {
        void Info(string message);
        void Warn(string message);
        void Error(string message, Exception ex = null);
        void Debug(string message);
    }

    internal sealed class SimpleFileLogger : ILog, IDisposable
    {
        private readonly string _logPath;
        private readonly object _gate = new object();
        private readonly Encoding _utf8 = new UTF8Encoding(false);
        private bool _disposed;

        public SimpleFileLogger(string logPath)
        {
            var dir = Path.GetDirectoryName(logPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            _logPath = logPath;
        }

        public void Info(string message) => Write("INFO", message);
        public void Warn(string message) => Write("WARN", message);
        public void Debug(string message) => Write("DEBUG", message);

        public void Error(string message, Exception ex = null)
        {
            Write("ERROR", ex == null ? message : message + Environment.NewLine + ex);
        }

        private void Write(string level, string message)
        {
            if (_disposed) return;
            var line = DateTime.UtcNow.ToString("O") + " [" + level + "] " + message + Environment.NewLine;
            lock (_gate)
            {
                try { File.AppendAllText(_logPath, line, _utf8); }
                catch { /* don’t crash on log failures */ }
            }
        }

        public void Dispose() { _disposed = true; }
    }
}

using System;
using System.Collections.Concurrent;
using System.IO;
using System.Text;
using System.Threading;


namespace Bs.SolidWorks.Tools.Logging {
    public enum LogLevel { Debug, Info, Warn, Error }
    public class Logger {
        TextWriter[] _writers;
        public Logger(params TextWriter[] writers) {
            _writers = writers ?? new TextWriter[] { Console.Out };
        }
    }

    public class LightLogger : IDisposable {
        readonly BlockingCollection<string> _queue = new BlockingCollection<string>(new ConcurrentQueue<string>());
        readonly Thread _worker;
        readonly string _folder;
        readonly string _baseFileName;
        readonly long _maxFileBytes;
        readonly int _maxFiles;
        TextWriter _writer;
        TextWriter? _externalWriter;
        string _currentPath;
        long _currentSize;
        bool _disposed;

        public LogLevel Level { get; set; } = LogLevel.Info;

        public LightLogger(string folder, string baseFileName = "Bs.SolidWorks.Tools.log", long maxFileBytes = 10 * 1024 * 1024, int maxFiles = 5, TextWriter? externalWriter = null) {
            _folder = folder ?? throw new ArgumentNullException(nameof(folder));
            _baseFileName = baseFileName ?? "app.log";
            _maxFileBytes = Math.Max(1024 * 1024, maxFileBytes);
            _maxFiles = Math.Max(1, maxFiles);

            Directory.CreateDirectory(_folder);
            _currentPath = Path.Combine(_folder, _baseFileName);
            _writer = RotateIfNeeded(true);

            _worker = new Thread(ProcessQueue) { IsBackground = true, Name = "LightLoggerWorker" };
            _worker.Start();
            _externalWriter = externalWriter;
        }

        TextWriter RotateIfNeeded(bool force = false) {
            if (!force && _writer != null && _currentSize < _maxFileBytes)
                return _writer;

            _writer?.Flush();
            _writer?.Dispose();

            // rotate files: app.log -> app.log.1, app.log.1 -> app.log.2 ...
            for (int i = _maxFiles - 1; i >= 1; i--) {
                string src = Path.Combine(_folder, _baseFileName + (i == 1 ? "" : $".{i - 1}"));
                string dst = Path.Combine(_folder, _baseFileName + $".{i}");
                if (File.Exists(src)) {
                    if (File.Exists(dst))
                        File.Delete(dst);
                    File.Move(src, dst);
                }
            }
            // current becomes .1
            string cur = Path.Combine(_folder, _baseFileName);
            string first = Path.Combine(_folder, _baseFileName + ".1");
            if (File.Exists(cur)) {
                if (File.Exists(first))
                    File.Delete(first);
                File.Move(cur, first);
            }

            TextWriter writer = new StreamWriter(new FileStream(cur, FileMode.Create, FileAccess.Write, FileShare.Read), Encoding.UTF8) { AutoFlush = true };
            _currentSize = 0;
            return writer;
        }

        string Format(LogLevel level, string msg) {
            return $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {msg}";
        }

        public void Log(LogLevel level, string message) {
            if (_disposed)
                return;
            if (level < Level)
                return;
            var line = Format(level, message);
            // best-effort enqueue
            try { _queue.Add(line); }
            catch { /* ignore */ }
        }

        public void Debug(string message) => Log(LogLevel.Debug, message);
        public void Info(string message) => Log(LogLevel.Info, message);
        public void Warn(string message) => Log(LogLevel.Warn, message);
        public void Error(string message) => Log(LogLevel.Error, message);

        void ProcessQueue() {
            try {
                foreach (var line in _queue.GetConsumingEnumerable()) {
                    try {
                        if(_externalWriter != null) {
                            _externalWriter.WriteLine(line);
                        }
                        var bytes = Encoding.UTF8.GetByteCount(line + Environment.NewLine);
                        _writer.WriteLine(line);
                        _currentSize += bytes;
                        if (_currentSize >= _maxFileBytes)
                            _writer = RotateIfNeeded(true);
                    }
                    catch {
                        // swallow logging exceptions to avoid crashing app
                    }
                }
            }
            catch {
                // worker exiting
            }
            finally {
                try { _writer?.Flush(); _writer?.Dispose(); }
                catch { }
            }
        }

        public void Flush() {
            // дождаться опустошения очереди
            while (_queue.Count > 0) {
                Thread.Sleep(10);
            }
            try { _writer?.Flush(); }
            catch { }
        }

        public void Dispose() {
            if (_disposed)
                return;
            _disposed = true;
            _queue.CompleteAdding();
            if (!_worker.Join(2000)) {
                try { _worker.Abort(); }
                catch { }
            }
            _queue.Dispose();
            try { _writer?.Flush(); _writer?.Dispose(); }
            catch { }
        }
    }
}

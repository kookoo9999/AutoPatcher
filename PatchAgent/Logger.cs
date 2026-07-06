using System;
using System.IO;

namespace PatchAgent
{
    internal class Logger
    {
        private readonly string _logFilePath;

        public Logger(string baseDirectory)
        {
            string logDir = Path.Combine(baseDirectory, "Logs");
            Directory.CreateDirectory(logDir);
            _logFilePath = Path.Combine(logDir, $"PatchAgent_{DateTime.Now:yyyyMMdd}.txt");
        }

        public void Info(string message) => Write("INFO", message);
        public void Warn(string message) => Write("WARN", message);
        public void Error(string message) => Write("ERROR", message);

        private void Write(string level, string message)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
            Console.WriteLine(line);
            try
            {
                File.AppendAllText(_logFilePath, line + Environment.NewLine);
            }
            catch
            {
                // 로깅 실패는 무시 (다음 사이클에서 재시도)
            }
        }
    }
}

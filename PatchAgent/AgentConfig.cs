using System;
using System.Collections.Generic;
using System.IO;

namespace PatchAgent
{
    internal class AgentConfig
    {
        public string CentralServer { get; private set; }
        public string Mode { get; private set; }
        public string PCType { get; private set; }
        public string ProcessName { get; private set; }
        public string InstallRoot { get; private set; }
        public int MaxJitterSeconds { get; private set; }
        public int BackupRetentionDays { get; private set; }

        public static AgentConfig Load(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"설정 파일을 찾을 수 없습니다: {path}");

            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string rawLine in File.ReadAllLines(path))
            {
                string line = rawLine.Trim();
                if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";")) continue;

                int idx = line.IndexOf('=');
                if (idx <= 0) continue;

                string key = line.Substring(0, idx).Trim();
                string value = line.Substring(idx + 1).Trim();
                values[key] = value;
            }

            var config = new AgentConfig
            {
                CentralServer = GetRequired(values, "CentralServer"),
                Mode = GetRequired(values, "Mode"),
                PCType = GetRequired(values, "PCType"),
                ProcessName = GetRequired(values, "ProcessName"),
                InstallRoot = GetRequired(values, "InstallRoot"),
                MaxJitterSeconds = GetOptionalInt(values, "MaxJitterSeconds", defaultValue: 60),
                BackupRetentionDays = GetOptionalInt(values, "BackupRetentionDays", defaultValue: 14),
            };

            return config;
        }

        private static string GetRequired(Dictionary<string, string> values, string key)
        {
            if (!values.TryGetValue(key, out string value) || string.IsNullOrWhiteSpace(value))
                throw new InvalidOperationException($"설정 항목 누락: {key}");
            return value;
        }

        private static int GetOptionalInt(Dictionary<string, string> values, string key, int defaultValue)
        {
            if (!values.TryGetValue(key, out string value) || string.IsNullOrWhiteSpace(value))
                return defaultValue;
            return int.TryParse(value, out int parsed) ? parsed : defaultValue;
        }
    }
}

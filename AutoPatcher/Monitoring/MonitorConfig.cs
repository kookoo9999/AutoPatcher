using System;
using System.Collections.Generic;
using System.IO;

namespace AutoPatcher
{
    /// <summary>exe 옆 AutoPatcher.ini. 사이트마다 서버가 다르므로 중앙 경로는 설치된 서버 기준 하나만 둔다.</summary>
    internal class MonitorConfig
    {
        public string CentralServer { get; set; } = string.Empty;
        public string TaskNameMain { get; set; } = "PatchAgent";
        public string TaskNameVision { get; set; } = "PatchAgent";
        public int StaleMinutes { get; set; } = 10;

        public static string FilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AutoPatcher.ini");

        public static MonitorConfig Load()
        {
            var config = new MonitorConfig();
            if (!File.Exists(FilePath))
            {
                try { config.Save(); }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) { }
                return config;
            }

            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string raw in File.ReadAllLines(FilePath))
            {
                string line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";")) continue;
                int idx = line.IndexOf('=');
                if (idx <= 0) continue;
                values[line.Substring(0, idx).Trim()] = line.Substring(idx + 1).Trim();
            }

            if (values.TryGetValue("CentralServer", out string central)) config.CentralServer = central;
            if (values.TryGetValue("TaskName_Main", out string tMain) && tMain.Length > 0) config.TaskNameMain = tMain;
            if (values.TryGetValue("TaskName_Vision", out string tVision) && tVision.Length > 0) config.TaskNameVision = tVision;
            if (values.TryGetValue("StaleMinutes", out string stale) && int.TryParse(stale, out int minutes) && minutes > 0)
                config.StaleMinutes = minutes;
            return config;
        }

        public void Save()
        {
            string content =
                "# AutoPatcher 모니터링 설정\r\n" +
                "# 중앙 배포 공유 폴더 (설비 PatchAgent.ini의 CentralServer와 같은 위치)\r\n" +
                $"CentralServer={CentralServer}\r\n" +
                "# 설비 PC에 등록된 PatchAgent 작업 스케줄러 이름 (즉시 적용 시 schtasks /run /tn 값)\r\n" +
                $"TaskName_Main={TaskNameMain}\r\n" +
                $"TaskName_Vision={TaskNameVision}\r\n" +
                "# 상태파일이 이 시간(분) 이상 갱신되지 않으면 무응답으로 표시\r\n" +
                $"StaleMinutes={StaleMinutes}\r\n";
            File.WriteAllText(FilePath, content);
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoPatcher
{
    /// <summary>
    /// exe 옆 AutoPatcher.ini. 사이트마다 서버가 따로이므로 모드/엑셀/중앙 경로를 설치된 서버 기준으로 고정한다.
    /// </summary>
    internal class MonitorConfig
    {
        public string Mode { get; set; } = string.Empty;
        public string ExcelPath { get; set; } = @"E:\AutoPatch";
        public string CentralServer { get; set; } = @"E:\AutoPatch\Distribution";
        public int StaleMinutes { get; set; } = 10;
        public int RefreshSeconds { get; set; } = 60;

        // 너무 짧으면 중앙 공유를 불필요하게 자주 읽는다. 상태파일은 에이전트 폴링(수 분)마다만 바뀐다.
        private const int MinRefreshSeconds = 5;

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

            if (values.TryGetValue("Mode", out string mode)) config.Mode = mode.ToUpperInvariant();
            if (values.TryGetValue("ExcelPath", out string excel) && excel.Length > 0) config.ExcelPath = excel;
            if (values.TryGetValue("CentralServer", out string central) && central.Length > 0) config.CentralServer = central;
            if (values.TryGetValue("StaleMinutes", out string stale) && int.TryParse(stale, out int staleMin) && staleMin > 0)
                config.StaleMinutes = staleMin;
            if (values.TryGetValue("RefreshSeconds", out string refresh) && int.TryParse(refresh, out int refreshSec) && refreshSec > 0)
                config.RefreshSeconds = Math.Max(MinRefreshSeconds, refreshSec);
            else if (values.TryGetValue("RefreshMinutes", out string oldRefresh) && int.TryParse(oldRefresh, out int refreshMin) && refreshMin > 0)
                config.RefreshSeconds = refreshMin * 60; // 이전 버전 ini 호환

            // 이전 버전 ini에는 없는 항목이 있으면 현재 값으로 다시 써서 편집할 수 있게 한다.
            string[] keys = { "Mode", "ExcelPath", "CentralServer", "StaleMinutes", "RefreshSeconds" };
            if (keys.Any(k => !values.ContainsKey(k)))
            {
                try { config.Save(); }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) { }
            }
            return config;
        }

        public void Save()
        {
            string content =
                "# AutoPatcher 모니터링 설정 (프로그램 시작 시 자동 로드)\r\n" +
                "# 이 서버의 사이트: LF / SII / BGA / COB\r\n" +
                $"Mode={Mode}\r\n" +
                "# 설비 목록 엑셀 파일. 폴더를 지정하면 그 안의 .xlsx 중 가장 최근 수정된 파일을 사용\r\n" +
                $"ExcelPath={ExcelPath}\r\n" +
                "# 중앙 배포 폴더 (설비 PatchAgent.ini의 CentralServer와 같은 위치)\r\n" +
                $"CentralServer={CentralServer}\r\n" +
                "# 상태파일이 이 시간(분) 이상 갱신되지 않으면 무응답으로 표시\r\n" +
                $"StaleMinutes={StaleMinutes}\r\n" +
                $"# 자동 갱신 주기(초, 최소 {MinRefreshSeconds})\r\n" +
                $"RefreshSeconds={RefreshSeconds}\r\n";
            File.WriteAllText(FilePath, content);
        }

        /// <summary>ExcelPath가 폴더면 그 안의 최신 .xlsx를 고른다. 없으면 null.</summary>
        public string ResolveExcelFile()
        {
            if (File.Exists(ExcelPath)) return ExcelPath;
            if (!Directory.Exists(ExcelPath)) return null;

            return new DirectoryInfo(ExcelPath)
                .GetFiles("*.xls*")
                .Where(f => !f.Name.StartsWith("~$")) // 엑셀이 열려 있을 때 생기는 잠금 파일
                .OrderByDescending(f => f.LastWriteTime)
                .Select(f => f.FullName)
                .FirstOrDefault();
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace AutoPatcher
{
    /// <summary>
    /// 중앙 공유폴더의 version.txt와 PatchAgent 상태파일(status\*.txt)을 읽어 설비별 상태를 판정한다.
    /// 설비 PC에는 접속하지 않는다.
    /// </summary>
    internal static class VersionMonitor
    {
        public static string MainPcType(string mode) => mode == "SII" ? "Inline" : "Main";
        public static string VisionPcType(string mode) => mode == "SII" ? @"Inline\IS" : "Vision";

        public static MonitorSnapshot Load(string pcTypeRoot)
        {
            var snap = new MonitorSnapshot { Root = pcTypeRoot };
            try
            {
                string versionFile = Path.Combine(pcTypeRoot, "version.txt");
                if (!File.Exists(versionFile))
                {
                    snap.Error = $"중앙 version.txt 없음: {versionFile}";
                    return snap;
                }
                snap.ServerVersion = File.ReadAllText(versionFile).Trim();

                string statusDir = Path.Combine(pcTypeRoot, "status");
                if (!Directory.Exists(statusDir)) return snap;

                foreach (string file in Directory.GetFiles(statusDir, "*.txt"))
                {
                    if (!file.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) continue;

                    AgentStatus status = TryReadStatus(file);
                    if (status == null) continue;

                    foreach (string ip in status.IPs)
                    {
                        // 같은 IP가 여러 파일에 있으면(PC 이름 변경 등) 가장 최근 것을 쓴다.
                        if (!snap.ByIp.TryGetValue(ip, out AgentStatus existing) || existing.LastWriteUtc < status.LastWriteUtc)
                            snap.ByIp[ip] = status;
                    }
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                snap.Error = $"중앙 공유 접근 실패({pcTypeRoot}): {ex.Message}";
            }
            return snap;
        }

        private static AgentStatus TryReadStatus(string file)
        {
            // PatchAgent가 파일을 교체하는 순간과 겹치면 잠깐 실패할 수 있어 짧게 재시도한다.
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    DateTime lastWriteUtc = File.GetLastWriteTimeUtc(file);
                    string[] lines = File.ReadAllLines(file);
                    return Parse(lines, lastWriteUtc);
                }
                catch (IOException)
                {
                    Thread.Sleep(100);
                }
                catch (UnauthorizedAccessException)
                {
                    return null;
                }
            }
            return null;
        }

        internal static AgentStatus Parse(string[] lines, DateTime lastWriteUtc)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in lines)
            {
                int idx = line.IndexOf('=');
                if (idx <= 0) continue;
                values[line.Substring(0, idx).Trim()] = line.Substring(idx + 1).Trim();
            }

            return new AgentStatus
            {
                MachineName = Get(values, "MachineName"),
                PCType = Get(values, "PCType"),
                IPs = (Get(values, "IPs") ?? string.Empty)
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => s.Length > 0)
                    .ToArray(),
                LocalVersion = Get(values, "LocalVersion"),
                ServerVersion = Get(values, "ServerVersion"),
                State = Get(values, "State"),
                Error = Get(values, "Error"),
                LastWriteUtc = lastWriteUtc,
            };
        }

        private static string Get(Dictionary<string, string> values, string key) =>
            values.TryGetValue(key, out string v) ? v : null;

        /// <summary>
        /// 판정 우선순위: 확인불가 → 미설치 → 무응답 → 일치 → 실패 → 적용대기 → 불일치.
        /// 무응답이 일치보다 앞서는 이유: 버전이 맞아도 에이전트가 죽어 있으면 다음 배포를 못 받는다.
        /// </summary>
        public static void Evaluate(MonitorCell cell, MonitorSnapshot snap, DateTime nowUtc, TimeSpan staleAfter)
        {
            if (!cell.HasIp)
            {
                cell.State = MonitorState.None;
                cell.Text = null;
                cell.Detail = null;
                return;
            }

            if (snap.Error != null)
            {
                cell.State = MonitorState.None;
                cell.Text = "확인 불가";
                cell.Detail = snap.Error;
                return;
            }

            if (!snap.ByIp.TryGetValue(cell.Ip, out AgentStatus st))
            {
                cell.State = MonitorState.NotInstalled;
                cell.Text = "-";
                cell.Detail = $"상태파일 없음 (에이전트 미설치 또는 구버전)\n경로: {Path.Combine(snap.Root, "status")}";
                return;
            }

            TimeSpan age = nowUtc - st.LastWriteUtc;
            bool sameVersion = VersionEquals(st.LocalVersion, snap.ServerVersion);

            if (age > staleAfter) cell.State = MonitorState.NoResponse;
            else if (sameVersion) cell.State = MonitorState.UpToDate;
            else if (st.State == "Fail" || st.State == "StageFail") cell.State = MonitorState.Failed;
            else if (st.State == "Waiting" && VersionEquals(st.ServerVersion, snap.ServerVersion)) cell.State = MonitorState.Pending;
            else cell.State = MonitorState.Mismatch;

            cell.Text = string.IsNullOrEmpty(st.LocalVersion) ? "(버전 없음)" : st.LocalVersion;
            cell.Detail = BuildDetail(st, snap, age);
        }

        private static bool VersionEquals(string a, string b) =>
            !string.IsNullOrEmpty(a) && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

        private static string BuildDetail(AgentStatus st, MonitorSnapshot snap, TimeSpan age)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"로컬 버전 : {st.LocalVersion}");
            sb.AppendLine($"서버 버전 : {snap.ServerVersion}");
            sb.AppendLine($"에이전트 상태 : {st.State}");
            sb.AppendLine($"마지막 확인 : {st.LastWriteUtc.ToLocalTime():MM-dd HH:mm:ss} ({FormatAge(age)} 전)");
            sb.Append($"PC : {st.MachineName} ({string.Join(", ", st.IPs)})");
            if (!string.IsNullOrEmpty(st.Error)) sb.Append($"\n에러 : {st.Error}");
            return sb.ToString();
        }

        private static string FormatAge(TimeSpan age)
        {
            if (age < TimeSpan.Zero) age = TimeSpan.Zero;
            if (age.TotalMinutes < 1) return $"{(int)age.TotalSeconds}초";
            if (age.TotalHours < 1) return $"{(int)age.TotalMinutes}분";
            if (age.TotalDays < 1) return $"{(int)age.TotalHours}시간";
            return $"{(int)age.TotalDays}일";
        }
    }
}

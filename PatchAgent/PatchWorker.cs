using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Threading;

namespace PatchAgent
{
    /// <summary>
    /// 한 사이클(폴링 -&gt; 필요시 적용) 수행. 대상 프로그램은 절대 강제 종료하지 않고,
    /// 이미 종료돼 있을 때만 패치를 적용한다 (스케줄러가 주기적으로 재호출).
    /// </summary>
    internal class PatchWorker
    {
        private readonly AgentConfig _config;
        private readonly Logger _log;

        private readonly string _localBin;
        private readonly string _localConfigDir;
        private readonly string _localBackup;
        private readonly string _localTempUpdate;
        private readonly string _localVersionMarker;
        private readonly string _stagedVersionMarker;

        private readonly string _centralModeRoot;
        private readonly string _centralVersionFile;
        private readonly string _centralStatusDir;

        public PatchWorker(AgentConfig config, Logger log)
        {
            _config = config;
            _log = log;

            _localBin = Path.Combine(config.InstallRoot, "bin");
            _localConfigDir = Path.Combine(config.InstallRoot, "config");
            _localBackup = Path.Combine(config.InstallRoot, "backup");
            _localTempUpdate = Path.Combine(config.InstallRoot, "temp_update");
            _localVersionMarker = Path.Combine(_localBin, "patch_agent_version.txt");
            _stagedVersionMarker = Path.Combine(_localTempUpdate, "staged_version.txt");

            _centralModeRoot = Path.Combine(config.CentralServer, config.PCType);
            _centralVersionFile = Path.Combine(_centralModeRoot, "version.txt");
            _centralStatusDir = Path.Combine(_centralModeRoot, "status");
        }

        public void RunCycle()
        {
            string remoteVersion = null;
            try
            {
                if (!File.Exists(_centralVersionFile))
                {
                    _log.Warn($"중앙 버전 파일에 접근할 수 없습니다: {_centralVersionFile}");
                    return;
                }

                remoteVersion = File.ReadAllText(_centralVersionFile).Trim();
                string localVersion = ReadLocalVersion();

                if (string.Equals(remoteVersion, localVersion, StringComparison.OrdinalIgnoreCase))
                {
                    _log.Info($"이미 최신 버전입니다 서버 : {remoteVersion} 로컬 : {localVersion}. 종료.");
                    WriteStatus("UpToDate", remoteVersion, null);
                    return;
                }

                EnsureStaged(remoteVersion);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                _log.Warn($"중앙 공유({_centralModeRoot}) 접근 실패 - 다음 주기에 재시도합니다: {ex.Message}");
                WriteStatus("StageFail", remoteVersion, ex.Message);
                return;
            }

            string processBaseName = Path.GetFileNameWithoutExtension(_config.ProcessName);
            //bool running = Process.GetProcessesByName(processBaseName).Any();
            int cnt = 0;
            while(Process.GetProcessesByName(processBaseName).Any())
            {
                if (cnt >= 10)
                {
                    _log.Info($"{_config.ProcessName} 실행 중 - 다음 주기에 재시도");
                    WriteStatus("Waiting", remoteVersion, null);
                    return;
                }
                Thread.Sleep(1000);
                cnt++;
            }

            try
            {
                _log.Info($"{_config.ProcessName} 종료 확인됨. 패치 적용 시작 (target version={remoteVersion}).");

                BackupDirectory(_localBin);
                if (string.Equals(_config.ProcessName, "HDSInspector.exe", StringComparison.OrdinalIgnoreCase))
                {
                    BackupDirectory(_localConfigDir);
                }

                CopyDirectoryOverwrite(_localTempUpdate, _localBin, excludeFileName: "staged_version.txt");

                Directory.CreateDirectory(_localBin);
                File.WriteAllText(_localVersionMarker, remoteVersion);

                if (string.Equals(_config.ProcessName, "HDSInspector.exe", StringComparison.OrdinalIgnoreCase))
                {
                    string exePath = Path.Combine(_localBin, _config.ProcessName);
                    if (File.Exists(exePath)) Process.Start(exePath);
                }

                ClearFolder(_localTempUpdate);

                CleanupOldBackups();

                _log.Info($"패치 적용 완료 (version={remoteVersion}).");
                WriteStatus("Success", remoteVersion, null);
            }
            catch (Exception ex)
            {
                _log.Error($"패치 적용 실패: {ex}");
                WriteStatus("Fail", remoteVersion, ex.Message);
            }
        }

        private void EnsureStaged(string remoteVersion)
        {
            string stagedVersion = File.Exists(_stagedVersionMarker)
                ? File.ReadAllText(_stagedVersionMarker).Trim()
                : null;

            if (string.Equals(stagedVersion, remoteVersion, StringComparison.OrdinalIgnoreCase))
                return;

            // 실제로 새 버전을 복사해야 할 때만 대기한다 - 다수 PC가 신버전 발견 순간에
            // 한꺼번에 central share로 몰리는 것을 막기 위함. 이미 스테이징된 버전을
            // 그대로 쓰는 사이클(=프로세스 종료 감지 -> 적용)에는 지연을 주지 않는다.
            ApplyJitter();

            string centralPackageDir = GetCentralPackageDir(remoteVersion);
            _log.Info($"새 버전 스테이징 중: {remoteVersion} (경로: {centralPackageDir})");
            ClearFolder(_localTempUpdate);
            try
            {
                Directory.CreateDirectory(_localTempUpdate);
                CopyDirectoryOverwrite(centralPackageDir, _localTempUpdate, excludeFileName: null);
                File.WriteAllText(_stagedVersionMarker, remoteVersion);
                _log.Info($"스테이징 완료: {remoteVersion}");
            }
            catch
            {
                // 빈/부분 temp_update가 남으면 설비 SW가 업데이트 대기로 오인하므로 정리한다.
                ClearFolder(_localTempUpdate);
                throw;
            }
        }

        private string ReadLocalVersion() =>
            File.Exists(_localVersionMarker) ? File.ReadAllText(_localVersionMarker).Trim() : string.Empty;

        // 배포측 계약: version.txt를 갱신하기 전에 package_{version} 폴더를 완전히 다 올려야 한다.
        // 버전별로 폴더가 분리되어 있어, 배포 도중인 다음 버전과 절대 경로가 겹치지 않는다
        // (= 폴링 중인 PC가 신/구 파일이 뒤섞인 내용을 읽을 수 없다).
        private string GetCentralPackageDir(string version) =>
            Path.Combine(_centralModeRoot, "package_" + version);

        private void ApplyJitter()
        {
            if (_config.MaxJitterSeconds <= 0) return;

            int waitSeconds = new Random().Next(0, _config.MaxJitterSeconds + 1);
            if (waitSeconds <= 0) return;

            _log.Info($"다수 PC 동시 접속 분산을 위해 {waitSeconds}초 대기.");
            Thread.Sleep(TimeSpan.FromSeconds(waitSeconds));
        }

        private static void ClearFolder(string path)
        {
            if (!Directory.Exists(path)) return;
            Directory.Delete(path, recursive: true);
        }

        private static void CopyDirectoryOverwrite(string sourceDir, string destinationDir, string excludeFileName)
        {
            var dir = new DirectoryInfo(sourceDir);
            if (!dir.Exists)
                throw new DirectoryNotFoundException($"소스 폴더를 찾을 수 없습니다: {sourceDir}");

            Directory.CreateDirectory(destinationDir);

            foreach (FileInfo file in dir.GetFiles())
            {
                if (excludeFileName != null && file.Name.Equals(excludeFileName, StringComparison.OrdinalIgnoreCase))
                    continue;

                string targetFilePath = Path.Combine(destinationDir, file.Name);
                file.CopyTo(targetFilePath, overwrite: true);
            }

            foreach (DirectoryInfo subDir in dir.GetDirectories())
            {
                if (subDir.Name.IndexOf(".git", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                string newDestinationDir = Path.Combine(destinationDir, subDir.Name);
                CopyDirectoryOverwrite(subDir.FullName, newDestinationDir, excludeFileName);
            }
        }

        private void BackupDirectory(string sourceDir)
        {
            if (!Directory.Exists(sourceDir)) return;
            if (!new DirectoryInfo(sourceDir).GetFiles().Any()) return;

            string backupType = sourceDir.Equals(_localConfigDir, StringComparison.OrdinalIgnoreCase) ? "config" : "bin";
            string backupPath = Path.Combine(_localBackup, DateTime.Now.ToString("yyMMdd"), backupType);
            Directory.CreateDirectory(backupPath);
            CopyDirectoryOverwrite(sourceDir, backupPath, excludeFileName: null);
        }

        // backup\yyMMdd\ 형식의 날짜 폴더 중 보존기간(기본 14일)이 지난 것을 삭제한다.
        // 패치마다 bin/config가 통째로 쌓이므로 정리하지 않으면 디스크가 무한 증가한다.
        private void CleanupOldBackups()
        {
            int retentionDays = _config.BackupRetentionDays;
            if (retentionDays <= 0) return; // 0 이하이면 정리 비활성화
            if (!Directory.Exists(_localBackup)) return;

            DateTime cutoff = DateTime.Now.Date.AddDays(-retentionDays);

            foreach (string dir in Directory.GetDirectories(_localBackup))
            {
                string name = Path.GetFileName(dir);
                // yyMMdd 형식이 아닌 폴더는 건드리지 않는다 (안전장치).
                if (!DateTime.TryParseExact(name, "yyMMdd", CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out DateTime folderDate))
                    continue;

                if (folderDate >= cutoff) continue;

                try
                {
                    Directory.Delete(dir, recursive: true);
                    _log.Info($"오래된 백업 삭제: {name} (보존 {retentionDays}일 초과).");
                }
                catch (Exception ex)
                {
                    _log.Warn($"백업 삭제 실패 ({name}): {ex.Message}");
                }
            }
        }

        // 매 사이클 기록(heartbeat). AutoPatcher 모니터링은 IPs로 엑셀의 설비와 매칭하고,
        // 파일 수정 시각으로 무응답을 판정한다.
        // State: UpToDate / Waiting / Success / Fail / StageFail
        private void WriteStatus(string state, string serverVersion, string errorMessage)
        {
            try
            {
                Directory.CreateDirectory(_centralStatusDir);
                string safePcType = _config.PCType.Replace('\\', '_').Replace('/', '_');
                string statusFile = Path.Combine(_centralStatusDir, $"{Environment.MachineName}_{safePcType}.txt");

                string localVersion;
                try { localVersion = ReadLocalVersion(); }
                catch (IOException) { localVersion = string.Empty; }

                string content =
                    $"MachineName={Environment.MachineName}\r\n" +
                    $"Mode={_config.Mode}\r\n" +
                    $"PCType={_config.PCType}\r\n" +
                    $"IPs={string.Join(",", GetLocalIPv4Addresses())}\r\n" +
                    $"LocalVersion={localVersion}\r\n" +
                    $"ServerVersion={serverVersion}\r\n" +
                    $"State={state}\r\n" +
                    $"Timestamp={DateTime.Now:yyyy-MM-dd HH:mm:ss}\r\n" +
                    (errorMessage != null ? $"Error={errorMessage.Replace("\r", " ").Replace("\n", " ")}\r\n" : string.Empty);
                WriteFileAtomic(statusFile, content);
            }
            catch (Exception ex)
            {
                _log.Warn($"중앙 상태 파일 기록 실패: {ex.Message}");
            }
        }

        // 모니터링 쪽이 쓰다 만 파일을 읽지 않도록 임시파일에 쓴 뒤 교체한다.
        private static void WriteFileAtomic(string path, string content)
        {
            string tmp = path + ".tmp";
            File.WriteAllText(tmp, content);
            try
            {
                if (File.Exists(path)) File.Replace(tmp, path, null);
                else File.Move(tmp, path);
            }
            catch (IOException)
            {
                // 일부 공유(SMB) 환경에서 Replace가 지원되지 않을 때의 대체 경로.
                File.Copy(tmp, path, overwrite: true);
                File.Delete(tmp);
            }
        }

        private static string[] GetLocalIPv4Addresses()
        {
            try
            {
                return NetworkInterface.GetAllNetworkInterfaces()
                    .Where(n => n.OperationalStatus == OperationalStatus.Up &&
                                n.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                    .SelectMany(n => n.GetIPProperties().UnicastAddresses)
                    .Where(a => a.Address.AddressFamily == AddressFamily.InterNetwork)
                    .Select(a => a.Address.ToString())
                    .Distinct()
                    .ToArray();
            }
            catch (NetworkInformationException)
            {
                return new string[0];
            }
        }
    }
}

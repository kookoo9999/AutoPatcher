using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

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
        private readonly string _centralPackageDir;
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

            _centralModeRoot = Path.Combine(config.CentralServer, config.Mode, config.PCType);
            _centralVersionFile = Path.Combine(_centralModeRoot, "version.txt");
            _centralPackageDir = Path.Combine(_centralModeRoot, "package");
            _centralStatusDir = Path.Combine(_centralModeRoot, "status");
        }

        public void RunCycle()
        {
            if (!File.Exists(_centralVersionFile))
            {
                _log.Warn($"중앙 버전 파일에 접근할 수 없습니다: {_centralVersionFile}");
                return;
            }

            string remoteVersion = File.ReadAllText(_centralVersionFile).Trim();
            string localVersion = File.Exists(_localVersionMarker)
                ? File.ReadAllText(_localVersionMarker).Trim()
                : string.Empty;

            if (string.Equals(remoteVersion, localVersion, StringComparison.OrdinalIgnoreCase))
            {
                _log.Info($"이미 최신 버전입니다 ({localVersion}). 종료.");
                return;
            }

            EnsureStaged(remoteVersion);

            string processBaseName = Path.GetFileNameWithoutExtension(_config.ProcessName);
            bool running = Process.GetProcessesByName(processBaseName).Any();

            if (running)
            {
                _log.Info($"{_config.ProcessName} 실행 중 - 종료될 때까지 대기 (다음 주기에 재시도).");
                WriteStatus("Waiting", remoteVersion, null);
                return;
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

            _log.Info($"새 버전 스테이징 중: {remoteVersion}");
            ClearFolder(_localTempUpdate);
            Directory.CreateDirectory(_localTempUpdate);
            CopyDirectoryOverwrite(_centralPackageDir, _localTempUpdate, excludeFileName: null);
            File.WriteAllText(_stagedVersionMarker, remoteVersion);
            _log.Info($"스테이징 완료: {remoteVersion}");
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

        private void WriteStatus(string result, string version, string errorMessage)
        {
            try
            {
                Directory.CreateDirectory(_centralStatusDir);
                string statusFile = Path.Combine(_centralStatusDir, $"{Environment.MachineName}_{_config.PCType}.txt");
                string content =
                    $"MachineName={Environment.MachineName}\r\n" +
                    $"Mode={_config.Mode}\r\n" +
                    $"PCType={_config.PCType}\r\n" +
                    $"Version={version}\r\n" +
                    $"Result={result}\r\n" +
                    $"Timestamp={DateTime.Now:yyyy-MM-dd HH:mm:ss}\r\n" +
                    (errorMessage != null ? $"Error={errorMessage}\r\n" : string.Empty);
                File.WriteAllText(statusFile, content);
            }
            catch (Exception ex)
            {
                _log.Warn($"중앙 상태 파일 기록 실패: {ex.Message}");
            }
        }
    }
}

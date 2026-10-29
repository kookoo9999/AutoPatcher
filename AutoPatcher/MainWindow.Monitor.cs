using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Resources;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace AutoPatcher
{
    /// <summary>모니터링 탭: 중앙 상태파일 기반 버전 모니터링, 배포 등록 및 에이전트 배포.</summary>
    public partial class MainWindow
    {
        private MonitorConfig _monConfig;
        private DispatcherTimer _monitorTimer;
        private bool _monitorRefreshing;
        private bool _monitorRefreshPending;
        private bool _monitorBusy;
        private string _startupMessage = string.Empty;
        private string _refreshError = string.Empty;
        private string _agentDeployStatus = string.Empty;
        private string _lastAgentSourceDir = string.Empty;

        private void InitMonitor()
        {
            _monConfig = MonitorConfig.Load();
            MonitorGrid.ItemsSource = DataGridItems;
            lblMonCentral.Text = $"중앙: {_monConfig.CentralServer}";

            _monitorTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(_monConfig.RefreshSeconds) };
            _monitorTimer.Tick += async (s, e) => await RefreshMonitorAsync();
            _monitorTimer.Start();

            // 엑셀(Interop) 로딩이 느려서 창이 먼저 뜬 뒤에 불러온다.
            Dispatcher.BeginInvoke(new Action(StartupLoad), DispatcherPriority.ApplicationIdle);
        }

        /// <summary>ini의 모드/엑셀을 기존 패치 탭과 같은 경로(모드 라디오 → 엑셀 로딩)로 적용한다.</summary>
        private void StartupLoad()
        {
            var radios = new Dictionary<string, RadioButton>
            {
                { "LF", rdLF }, { "SII", rdSII }, { "BGA", rdBGA }, { "COB", rdCOB },
            };
            if (!radios.TryGetValue(_monConfig.Mode, out RadioButton modeRadio))
            {
                SetStartupMessage($"AutoPatcher.ini의 Mode를 설정하세요 (LF/SII/BGA/COB): {MonitorConfig.FilePath}");
                return;
            }

            string excel = _monConfig.ResolveExcelFile();
            if (excel == null)
                SetStartupMessage($"설비 목록 엑셀을 찾을 수 없습니다: {_monConfig.ExcelPath}");
            else
            {
                ExcelPath = excel;
                lblCurExcel.Content = excel;
            }

            // Radiobutton_Checked가 엑셀을 읽고 OnMachineListLoaded를 호출한다.
            modeRadio.IsChecked = true;
        }

        /// <summary>엑셀 로딩/모드 변경 후 호출.</summary>
        private void OnMachineListLoaded()
        {
            if (_monConfig == null) return;

            lblMonTitle.Text = string.IsNullOrEmpty(ModeType) ? "버전 모니터링" : $"{ModeType} 버전 모니터링";
            lblMonExcel.Text = string.IsNullOrEmpty(ExcelPath) ? "엑셀: -" : $"엑셀: {ExcelPath} ({ModeType} 시트)";

            _ = RefreshMonitorAsync();
        }

        private static MonitorCell[] CellsOf(RowData row) =>
            new[] { row.M1, row.M2, row.M3, row.M4, row.M5, row.M6 };

        private void SetStartupMessage(string message)
        {
            _startupMessage = message;
            Log($"[Monitor] {message}", LogLevel.WARN);
            ShowMonitorMessage();
        }

        private void SetRefreshError(string error)
        {
            // 같은 오류를 매 주기 로그에 반복해서 남기지 않는다.
            if (error == _refreshError) return;
            _refreshError = error;
            if (!string.IsNullOrEmpty(error)) Log($"[Monitor] {error}", LogLevel.WARN);
            ShowMonitorMessage();
        }

        private void SetAgentDeployStatus(string status)
        {
            _agentDeployStatus = status ?? string.Empty;
            ShowMonitorMessage();
        }

        private void ShowMonitorMessage() =>
            lblMonMessage.Text = string.Join("  /  ", new[] { _agentDeployStatus, _startupMessage, _refreshError }.Where(m => !string.IsNullOrEmpty(m)));

        private void SetMonitorBusy(bool busy)
        {
            _monitorBusy = busy;
            btnMonDeploy.IsEnabled = !busy;
            btnMonAgentDeploy.IsEnabled = !busy;
        }

        #region 새로고침

        private async Task RefreshMonitorAsync()
        {
            if (_monitorRefreshing)
            {
                // 진행 중인 갱신이 끝나면 한 번 더 돈다 (모드 변경 직후 요청이 묻히지 않도록).
                _monitorRefreshPending = true;
                return;
            }
            if (string.IsNullOrEmpty(ModeType)) return;

            _monitorRefreshing = true;
            btnMonRefresh.IsEnabled = false;
            try
            {
                string mode = ModeType;
                string central = _monConfig.CentralServer;

                MonitorSnapshot main = null, vision = null;
                await Task.Run(() =>
                {
                    main = VersionMonitor.Load(Path.Combine(central, VersionMonitor.MainPcType(mode)));
                    vision = VersionMonitor.Load(Path.Combine(central, VersionMonitor.VisionPcType(mode)));
                });

                // 갱신 도중 모드가 바뀌었으면 이 결과는 버리고 새 모드로 다시 돈다.
                if (mode != ModeType)
                {
                    _monitorRefreshPending = true;
                }
                else
                {
                    ApplySnapshots(main, vision);
                    lblMonUpdated.Text = $"자동 갱신 {_monConfig.RefreshSeconds}초 · 마지막 갱신 {DateTime.Now:HH:mm:ss}";

                    SetRefreshError(string.Join("  /  ", new[] { main.Error, vision.Error }.Where(e => e != null)));
                }
            }
            catch (Exception ex)
            {
                SetRefreshError($"갱신 실패: {ex.Message}");
            }
            finally
            {
                _monitorRefreshing = false;
                btnMonRefresh.IsEnabled = true;
            }

            if (_monitorRefreshPending)
            {
                _monitorRefreshPending = false;
                await RefreshMonitorAsync();
            }
        }

        private void ApplySnapshots(MonitorSnapshot main, MonitorSnapshot vision)
        {
            lblServerMain.Text = string.IsNullOrEmpty(main.ServerVersion) ? "-" : main.ServerVersion;
            lblServerVision.Text = string.IsNullOrEmpty(vision.ServerVersion) ? "-" : vision.ServerVersion;
            lblServerMainType.Text = $"중앙 {VersionMonitor.MainPcType(ModeType)}";
            lblServerVisionType.Text = $"중앙 {VersionMonitor.VisionPcType(ModeType)}";

            DateTime nowUtc = DateTime.UtcNow;
            TimeSpan stale = TimeSpan.FromMinutes(_monConfig.StaleMinutes);
            var counts = new Dictionary<MonitorState, int>();

            foreach (RowData row in DataGridItems)
            {
                MonitorCell[] cells = CellsOf(row);
                for (int i = 0; i < cells.Length; i++)
                {
                    VersionMonitor.Evaluate(cells[i], i == 0 ? main : vision, nowUtc, stale);
                    if (cells[i].HasIp)
                        counts[cells[i].State] = (counts.TryGetValue(cells[i].State, out int n) ? n : 0) + 1;
                }
            }

            cntUpToDate.Text = CountOf(counts, MonitorState.UpToDate);
            cntMismatch.Text = CountOf(counts, MonitorState.Mismatch);
            cntPending.Text = CountOf(counts, MonitorState.Pending);
            cntFailed.Text = CountOf(counts, MonitorState.Failed);
            cntNoResponse.Text = CountOf(counts, MonitorState.NoResponse);
            cntNotInstalled.Text = CountOf(counts, MonitorState.NotInstalled);
        }

        private static string CountOf(Dictionary<MonitorState, int> counts, MonitorState state) =>
            (counts.TryGetValue(state, out int n) ? n : 0).ToString();

        private async void btnMonRefresh_Click(object sender, RoutedEventArgs e) => await RefreshMonitorAsync();

        #endregion

        #region 배포 등록

        private async void btnMonDeploy_Click(object sender, RoutedEventArgs e)
        {
            if (_monitorBusy) return;
            if (string.IsNullOrEmpty(ModeType))
            {
                MessageBox.Show("모드가 설정되지 않았습니다. AutoPatcher.ini의 Mode를 확인하세요.", "배포 등록");
                return;
            }

            string mainType = VersionMonitor.MainPcType(ModeType);
            string visionType = VersionMonitor.VisionPcType(ModeType);
            var dialog = new DeployDialog(_monConfig.CentralServer, mainType, visionType, SourceDirectory) { Owner = this };
            if (dialog.ShowDialog() != true) return;

            string root = Path.Combine(_monConfig.CentralServer, dialog.TargetPcType);
            string confirm =
                $"[{dialog.TargetPcType}] {dialog.CurrentVersion} → {dialog.NewVersion}\n" +
                $"패치 폴더: {dialog.SourceDir}\n\n" +
                "등록하면 해당 타입의 모든 설비가 다음 폴링 때 새 버전을 받아갑니다. 진행할까요?";
            if (MessageBox.Show(confirm, "배포 등록", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;

            SetMonitorBusy(true);
            try
            {
                Log($"[Monitor] 배포 등록 시작: {dialog.TargetPcType} {dialog.NewVersion}");
                await Task.Run(() => PackageRegistrar.Register(root, dialog.NewVersion, dialog.SourceDir,
                    msg => Dispatcher.InvokeAsync(() => Log($"[Monitor] {msg}"))));
                Log($"[Monitor] 배포 등록 완료: {dialog.TargetPcType} {dialog.NewVersion}");
            }
            catch (Exception ex)
            {
                Log($"[Monitor] 배포 등록 실패: {ex.Message}", LogLevel.ERROR);
                MessageBox.Show($"배포 등록 실패\n{ex.Message}", "배포 등록", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetMonitorBusy(false);
            }
            await RefreshMonitorAsync();
        }

        #endregion

        #region 에이전트 배포

        private class AgentDeployTarget
        {
            public string Ip;
            public string UnitLabel;
            public bool IsVision;
        }

        private async void btnMonAgentDeploy_Click(object sender, RoutedEventArgs e)
        {
            if (_monitorBusy) return;
            if (string.IsNullOrEmpty(ModeType))
            {
                MessageBox.Show("모드가 설정되지 않았습니다. AutoPatcher.ini의 Mode를 확인하세요.", "에이전트 배포");
                return;
            }

            List<AgentDeployTarget> targets = CollectAgentDeployTargets();
            if (targets.Count == 0)
            {
                MessageBox.Show("배포할 설비 목록이 없습니다. 엑셀 파일이 정상적으로 로드되었는지 확인하세요.", "에이전트 배포");
                return;
            }

            var folderDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = "배포할 PatchAgent 폴더 선택 (설비별 PatchAgent.ini는 제외하고 덮어씁니다)",
            };
            if (!string.IsNullOrEmpty(_lastAgentSourceDir) && Directory.Exists(_lastAgentSourceDir))
                folderDialog.InitialDirectory = _lastAgentSourceDir;

            if (folderDialog.ShowDialog() != CommonFileDialogResult.Ok) return;

            string sourceDir = folderDialog.FileName.Trim().TrimEnd('\\');
            if (!Directory.Exists(sourceDir))
            {
                MessageBox.Show("선택한 폴더가 존재하지 않습니다.", "에이전트 배포", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int fileCount = CountAgentDeployFiles(sourceDir);
            if (fileCount == 0)
            {
                MessageBox.Show("선택한 폴더에 배포할 파일이 없습니다 (PatchAgent.ini 제외).", "에이전트 배포", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool hasAgentExe = File.Exists(Path.Combine(sourceDir, "PatchAgent.exe"));
            if (!hasAgentExe)
            {
                string warnExe =
                    $"선택한 폴더에 PatchAgent.exe가 없습니다.\n폴더: {sourceDir}\n\n" +
                    "그래도 이 폴더의 파일들을 에이전트 경로에 배포할까요?";
                if (MessageBox.Show(warnExe, "에이전트 배포 확인", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                    return;
            }

            _lastAgentSourceDir = sourceDir;

            string mode = ModeType;
            int mainCount = targets.Count(t => !t.IsVision);
            int visionCount = targets.Count(t => t.IsVision);
            string mainSubPath = $@"{VersionMonitor.MainPcType(mode)}\PatchAgent";
            string visionSubPath = $@"{VersionMonitor.VisionPcType(mode)}\PatchAgent";

            string confirmMsg =
                $"모드: {mode}\n" +
                $"배포 대상: 전 설비 총 {targets.Count}대 (Main {mainCount}대, Vision {visionCount}대)\n" +
                $"원본 폴더: {sourceDir} (파일 {fileCount}개)\n" +
                $"대상 경로:\n" +
                $"  - Main: \\\\{{IP}}\\(D\\){mainSubPath}\n" +
                $"  - Vision: \\\\{{IP}}\\(D\\){visionSubPath}\n\n" +
                "※ 각 설비의 기존 PatchAgent.ini와 Logs 폴더는 유지하고 나머지 파일을 덮어씁니다.\n" +
                "진행할까요?";
            if (MessageBox.Show(confirmMsg, "에이전트 배포", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            SetMonitorBusy(true);
            SetAgentDeployStatus($"에이전트 배포 준비 중... (0/{targets.Count})");
            Log($"[AgentDeploy] 에이전트 배포 시작: 모드={mode}, 대상={targets.Count}대 (Main {mainCount}, Vision {visionCount}), 원본={sourceDir}");

            int completed = 0;
            int succeeded = 0;
            var failures = new ConcurrentBag<string>();

            try
            {
                using (var semaphore = new SemaphoreSlim(10))
                {
                    var tasks = targets.Select(async target =>
                    {
                        await semaphore.WaitAsync();
                        try
                        {
                            (bool ok, string detail) = await Task.Run(() => DeployAgentToTarget(target, mode, sourceDir));
                            int done = Interlocked.Increment(ref completed);

                            if (ok)
                            {
                                Interlocked.Increment(ref succeeded);
                                _ = Dispatcher.InvokeAsync(() =>
                                {
                                    Log($"[AgentDeploy] [{target.UnitLabel} ({target.Ip})] 배포 완료 -> {detail}");
                                    SetAgentDeployStatus($"에이전트 배포 중... ({done}/{targets.Count})");
                                });
                            }
                            else
                            {
                                failures.Add($"{target.UnitLabel} ({target.Ip}): {detail}");
                                _ = Dispatcher.InvokeAsync(() =>
                                {
                                    Log($"[AgentDeploy] [{target.UnitLabel} ({target.Ip})] 배포 실패: {detail}", LogLevel.ERROR);
                                    SetAgentDeployStatus($"에이전트 배포 중... ({done}/{targets.Count})");
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            int done = Interlocked.Increment(ref completed);
                            failures.Add($"{target.UnitLabel} ({target.Ip}): {ex.Message}");
                            _ = Dispatcher.InvokeAsync(() =>
                            {
                                Log($"[AgentDeploy] [{target.UnitLabel} ({target.Ip})] 예외 발생: {ex.Message}", LogLevel.ERROR);
                                SetAgentDeployStatus($"에이전트 배포 중... ({done}/{targets.Count})");
                            });
                        }
                        finally
                        {
                            semaphore.Release();
                        }
                    }).ToList();

                    await Task.WhenAll(tasks);
                }

                int failCount = failures.Count;
                string summaryStatus = failCount == 0
                    ? $"에이전트 배포 완료 (성공 {succeeded}/{targets.Count})"
                    : $"에이전트 배포 완료 (성공 {succeeded}, 실패 {failCount})";
                SetAgentDeployStatus(summaryStatus);
                Log($"[AgentDeploy] {summaryStatus}", failCount == 0 ? LogLevel.INFO : LogLevel.WARN);

                string resultMessage = $"에이전트 배포가 완료되었습니다.\n\n성공: {succeeded}대 / 전체: {targets.Count}대 (실패: {failCount}대)";
                if (failCount > 0)
                {
                    var preview = failures.OrderBy(f => f).Take(15).ToList();
                    resultMessage += "\n\n[실패 목록]\n" + string.Join("\n", preview);
                    if (failCount > preview.Count)
                        resultMessage += $"\n... 외 {failCount - preview.Count}건 (상세 내역은 로그 확인)";
                }

                MessageBox.Show(
                    resultMessage,
                    "에이전트 배포 결과",
                    MessageBoxButton.OK,
                    failCount == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            finally
            {
                SetMonitorBusy(false);
            }
        }

        private List<AgentDeployTarget> CollectAgentDeployTargets()
        {
            var targets = new List<AgentDeployTarget>();
            var seenIps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (RowData row in DataGridItems)
            {
                string unit = string.IsNullOrWhiteSpace(row.InspectionUnit) ? "Unknown" : row.InspectionUnit.Trim();
                AddTargetIfValid(targets, seenIps, row.PC1, $"{unit}_MAIN", isVision: false);
                AddTargetIfValid(targets, seenIps, row.PC2, $"{unit}_V1", isVision: true);
                AddTargetIfValid(targets, seenIps, row.PC3, $"{unit}_V2", isVision: true);
                AddTargetIfValid(targets, seenIps, row.PC4, $"{unit}_V3", isVision: true);
                AddTargetIfValid(targets, seenIps, row.PC5, $"{unit}_V4", isVision: true);
                AddTargetIfValid(targets, seenIps, row.PC6, $"{unit}_V5", isVision: true);
            }

            return targets;
        }

        private static void AddTargetIfValid(
            List<AgentDeployTarget> targets,
            HashSet<string> seenIps,
            string ip,
            string unitLabel,
            bool isVision)
        {
            string trimmed = ip?.Trim();
            if (string.IsNullOrEmpty(trimmed)) return;
            if (!seenIps.Add(trimmed)) return;

            targets.Add(new AgentDeployTarget
            {
                Ip = trimmed,
                UnitLabel = unitLabel,
                IsVision = isVision,
            });
        }

        private (bool ok, string detail) DeployAgentToTarget(AgentDeployTarget target, string mode, string sourceDir)
        {
            if (!GetPingResult(target.Ip))
                return (false, "Ping 응답 없음");

            string strPCtype = target.IsVision
                ? VersionMonitor.VisionPcType(mode)
                : VersionMonitor.MainPcType(mode);

            string[] remotePathList =
            {
                $@"\\{target.Ip}\{strPCtype.ToLowerInvariant()}",
                $@"\\{target.Ip}\{strPCtype}",
                $@"\\{target.Ip}\D\{strPCtype.ToLowerInvariant()}",
                $@"\\{target.Ip}\D\{strPCtype}",
                $@"\\{target.Ip}\d\{strPCtype.ToLowerInvariant()}",
                $@"\\{target.Ip}\d\{strPCtype}",
            };

            ResourceManager rscManager = new ResourceManager("AutoPatcher.Resource.UserInfo", typeof(MainWindow).Assembly);
            string roleKey = target.IsVision ? "VISION" : "MAIN";
            string id = rscManager.GetString($"{mode.ToUpperInvariant()}_{roleKey}_ID");
            string pw = rscManager.GetString($"{mode.ToUpperInvariant()}_{roleKey}_PW") ?? "vision";

            string remoteRoot = null;
            foreach (string pathAttempt in remotePathList)
            {
                bool connected = false;
                if (!string.IsNullOrEmpty(id))
                    connected = ConnectNetworkDrive(pathAttempt, id, pw);

                if (!connected)
                {
                    connected = ConnectNetworkDrive(pathAttempt, "eav", pw);
                    if (!connected) ConnectNetworkDrive(pathAttempt, "slf", pw);
                }

                try
                {
                    if (Directory.Exists(pathAttempt))
                    {
                        remoteRoot = pathAttempt;
                        break;
                    }
                }
                catch
                {
                    // 다음 후보 경로 계속 시도
                }
            }

            if (string.IsNullOrEmpty(remoteRoot))
                return (false, $"원격 경로 접근 불가 ({strPCtype})");

            string targetAgentDir = Path.Combine(remoteRoot, "PatchAgent");
            try
            {
                CopyAgentDirectory(sourceDir, targetAgentDir);
                return (true, targetAgentDir);
            }
            catch (Exception ex)
            {
                return (false, $"파일 복사 실패 ({targetAgentDir}): {ex.Message}");
            }
        }

        private static int CountAgentDeployFiles(string sourceDir)
        {
            int count = 0;
            var dir = new DirectoryInfo(sourceDir);
            if (!dir.Exists) return 0;

            foreach (FileInfo file in dir.GetFiles())
            {
                if (!IsExcludedAgentFile(file.Name))
                    count++;
            }

            foreach (DirectoryInfo subDir in dir.GetDirectories())
            {
                if (IsExcludedAgentDir(subDir.Name)) continue;
                count += CountAgentDeployFiles(subDir.FullName);
            }

            return count;
        }

        private static void CopyAgentDirectory(string sourceDir, string destinationDir)
        {
            var dir = new DirectoryInfo(sourceDir);
            if (!dir.Exists)
                throw new DirectoryNotFoundException($"원본 폴더를 찾을 수 없습니다: {sourceDir}");

            Directory.CreateDirectory(destinationDir);

            foreach (FileInfo file in dir.GetFiles())
            {
                if (IsExcludedAgentFile(file.Name)) continue;

                string targetFilePath = Path.Combine(destinationDir, file.Name);
                CopyFileWithRetry(file.FullName, targetFilePath);
            }

            foreach (DirectoryInfo subDir in dir.GetDirectories())
            {
                if (IsExcludedAgentDir(subDir.Name)) continue;

                string nextDestDir = Path.Combine(destinationDir, subDir.Name);
                CopyAgentDirectory(subDir.FullName, nextDestDir);
            }
        }

        private static void CopyFileWithRetry(string sourceFile, string targetFile)
        {
            const int maxAttempts = 3;
            for (int attempt = 1; ; attempt++)
            {
                try
                {
                    File.Copy(sourceFile, targetFile, overwrite: true);
                    return;
                }
                catch (IOException) when (attempt < maxAttempts)
                {
                    // 설비에서 스케줄러로 PatchAgent.exe가 순간 실행 중일 수 있으므로 짧게 대기 후 재시도
                    Thread.Sleep(1500);
                }
            }
        }

        private static bool IsExcludedAgentFile(string fileName) =>
            fileName.Equals("PatchAgent.ini", StringComparison.OrdinalIgnoreCase);

        private static bool IsExcludedAgentDir(string dirName) =>
            dirName.IndexOf(".git", StringComparison.OrdinalIgnoreCase) >= 0 ||
            dirName.Equals("Logs", StringComparison.OrdinalIgnoreCase);

        #endregion
    }
}

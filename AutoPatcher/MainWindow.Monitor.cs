using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Resources;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace AutoPatcher
{
    /// <summary>모니터링 탭: 중앙 상태파일 기반 버전 모니터링, 배포 등록, 즉시 적용(schtasks /run).</summary>
    public partial class MainWindow
    {
        private MonitorConfig _monConfig;
        private ListCollectionView _monitorView;
        private DispatcherTimer _monitorTimer;
        private bool _monitorRefreshing;
        private bool _monitorRefreshPending;
        private bool _monitorBusy;

        private void InitMonitor()
        {
            _monConfig = MonitorConfig.Load();
            lblCentralPath.Text = string.IsNullOrWhiteSpace(_monConfig.CentralServer) ? "(미설정)" : _monConfig.CentralServer;

            _monitorView = new ListCollectionView(DataGridItems);
            _monitorView.Filter = MonitorRowFilter;
            MonitorGrid.ItemsSource = _monitorView;

            _monitorTimer = new DispatcherTimer();
            _monitorTimer.Tick += async (s, e) => await RefreshMonitorAsync();
            UpdateMonitorTimer();
        }

        /// <summary>엑셀 로딩/모드 변경 후 호출.</summary>
        private void OnMachineListLoaded()
        {
            if (_monConfig == null) return;
            lblMonExcel.Text = ExcelPath ?? string.Empty;
            _ = RefreshMonitorAsync();
        }

        private static MonitorCell[] CellsOf(RowData row) =>
            new[] { row.M1, row.M2, row.M3, row.M4, row.M5, row.M6 };

        private static bool IsMainColumn(int index) => index == 0;

        private bool MonitorRowFilter(object item)
        {
            if (chkMonIssuesOnly?.IsChecked != true) return true;
            var row = item as RowData;
            return row != null && CellsOf(row).Any(c => c.HasIp && c.State != MonitorState.UpToDate && c.State != MonitorState.None);
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

            if (string.IsNullOrEmpty(ModeType))
            {
                lblMonMessage.Text = "모드(LF/SII/BGA/COB)를 선택하세요.";
                return;
            }
            if (string.IsNullOrWhiteSpace(_monConfig.CentralServer))
            {
                lblMonMessage.Text = "중앙 경로가 설정되지 않았습니다. [변경]으로 지정하세요.";
                return;
            }

            _monitorRefreshing = true;
            btnMonRefresh.IsEnabled = false;
            try
            {
                string mode = ModeType;
                string central = _monConfig.CentralServer;
                lblMonMessage.Text = "갱신 중...";

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
                    lblMonUpdated.Text = DateTime.Now.ToString("HH:mm:ss");

                    var errors = new[] { main.Error, vision.Error }.Where(e => e != null).ToList();
                    lblMonMessage.Text = errors.Count == 0 ? string.Empty : string.Join("  /  ", errors);
                }
            }
            catch (Exception ex)
            {
                lblMonMessage.Text = $"갱신 실패: {ex.Message}";
                Log($"[Monitor] 갱신 실패: {ex.Message}", LogLevel.ERROR);
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
            lblServerMain.Text = main.ServerVersion ?? "-";
            lblServerVision.Text = vision.ServerVersion ?? "-";
            lblServerMainType.Text = VersionMonitor.MainPcType(ModeType);
            lblServerVisionType.Text = VersionMonitor.VisionPcType(ModeType);

            DateTime nowUtc = DateTime.UtcNow;
            TimeSpan stale = TimeSpan.FromMinutes(_monConfig.StaleMinutes);
            var counts = new Dictionary<MonitorState, int>();

            foreach (RowData row in DataGridItems)
            {
                MonitorCell[] cells = CellsOf(row);
                for (int i = 0; i < cells.Length; i++)
                {
                    VersionMonitor.Evaluate(cells[i], IsMainColumn(i) ? main : vision, nowUtc, stale);
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

            // 색상은 바인딩으로 바로 반영된다. 재필터링(스크롤 초기화 유발)은 필터 사용 중일 때만.
            if (chkMonIssuesOnly.IsChecked == true) _monitorView.Refresh();
        }

        private static string CountOf(Dictionary<MonitorState, int> counts, MonitorState state) =>
            (counts.TryGetValue(state, out int n) ? n : 0).ToString();

        private async void btnMonRefresh_Click(object sender, RoutedEventArgs e) => await RefreshMonitorAsync();

        private void chkMonAutoRefresh_Changed(object sender, RoutedEventArgs e) => UpdateMonitorTimer();

        private void cmbMonInterval_SelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateMonitorTimer();

        private void UpdateMonitorTimer()
        {
            if (_monitorTimer == null) return;
            _monitorTimer.Stop();
            if (chkMonAutoRefresh.IsChecked != true) return;

            int minutes = 1;
            if (cmbMonInterval.SelectedItem is ComboBoxItem item)
                int.TryParse(item.Content?.ToString(), out minutes);
            _monitorTimer.Interval = TimeSpan.FromMinutes(Math.Max(1, minutes));
            _monitorTimer.Start();
        }

        private void chkMonIssuesOnly_Changed(object sender, RoutedEventArgs e) => _monitorView?.Refresh();

        #endregion

        #region 설정 / 선택

        private async void btnCentralPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog { IsFolderPicker = true, Title = "중앙 배포 공유 폴더 선택 (PatchAgent.ini의 CentralServer)" };
            if (dialog.ShowDialog() != CommonFileDialogResult.Ok) return;

            _monConfig.CentralServer = dialog.FileName;
            try
            {
                _monConfig.Save();
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                Log($"[Monitor] 설정 저장 실패({MonitorConfig.FilePath}): {ex.Message}", LogLevel.WARN);
            }
            lblCentralPath.Text = _monConfig.CentralServer;
            Log($"[Monitor] 중앙 경로 설정: {_monConfig.CentralServer}");
            await RefreshMonitorAsync();
        }

        private void MonitorHeaderCheck_Click(object sender, RoutedEventArgs e)
        {
            var checkBox = (CheckBox)sender;
            int index = int.Parse((string)checkBox.Tag) - 1;
            bool value = checkBox.IsChecked == true;

            foreach (RowData row in _monitorView.Cast<RowData>())
            {
                MonitorCell cell = CellsOf(row)[index];
                if (cell.HasIp) cell.Selected = value;
            }
        }

        private void btnMonClearSelection_Click(object sender, RoutedEventArgs e)
        {
            foreach (RowData row in DataGridItems)
                foreach (MonitorCell cell in CellsOf(row))
                    cell.Selected = false;
        }

        #endregion

        #region 배포 등록

        private async void btnMonDeploy_Click(object sender, RoutedEventArgs e)
        {
            if (_monitorBusy) return;
            if (string.IsNullOrEmpty(ModeType) || string.IsNullOrWhiteSpace(_monConfig.CentralServer))
            {
                MessageBox.Show("모드와 중앙 경로를 먼저 설정하세요.", "배포 등록");
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

            _monitorBusy = true;
            SetMonitorActionsEnabled(false);
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
                _monitorBusy = false;
                SetMonitorActionsEnabled(true);
            }
            await RefreshMonitorAsync();
        }

        private void SetMonitorActionsEnabled(bool enabled)
        {
            btnMonDeploy.IsEnabled = enabled;
            btnMonRunNow.IsEnabled = enabled;
        }

        #endregion

        #region 즉시 적용 (원격 PatchAgent 작업 실행)

        private async void btnMonRunNow_Click(object sender, RoutedEventArgs e)
        {
            if (_monitorBusy) return;
            if (string.IsNullOrEmpty(ModeType))
            {
                MessageBox.Show("모드를 먼저 선택하세요.", "즉시 적용");
                return;
            }

            // (IP, Main 여부) 단위. SII는 같은 IP라도 Main(Inline)/Vision(IS) 작업이 따로일 수 있다.
            var targets = new List<Tuple<string, bool>>();
            foreach (RowData row in DataGridItems)
            {
                MonitorCell[] cells = CellsOf(row);
                for (int i = 0; i < cells.Length; i++)
                {
                    var t = Tuple.Create(cells[i].Ip, IsMainColumn(i));
                    if (cells[i].HasIp && cells[i].Selected && !targets.Contains(t)) targets.Add(t);
                }
            }

            if (targets.Count == 0)
            {
                MessageBox.Show("선택된 설비가 없습니다.", "즉시 적용");
                return;
            }

            string confirm =
                $"선택한 {targets.Count}대에서 PatchAgent 작업을 즉시 실행합니다.\n" +
                "실행 중인 검사 프로그램은 종료하지 않으며, 종료된 설비만 바로 적용됩니다. 진행할까요?";
            if (MessageBox.Show(confirm, "즉시 적용", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;

            string mode = ModeType;
            _monitorBusy = true;
            SetMonitorActionsEnabled(false);
            int ok = 0, fail = 0;
            try
            {
                using (var gate = new SemaphoreSlim(10))
                {
                    var tasks = targets.Select(async t =>
                    {
                        await gate.WaitAsync();
                        try
                        {
                            bool success = await RunAgentTaskAsync(t.Item1, t.Item2, mode);
                            if (success) Interlocked.Increment(ref ok); else Interlocked.Increment(ref fail);
                        }
                        finally
                        {
                            gate.Release();
                        }
                    }).ToList();
                    await Task.WhenAll(tasks);
                }
                Log($"[Monitor] 즉시 적용 요청 완료: 성공 {ok}, 실패 {fail}. 상태는 에이전트 완료 후 새로고침에 반영됩니다.",
                    fail > 0 ? LogLevel.WARN : LogLevel.INFO);
            }
            finally
            {
                _monitorBusy = false;
                SetMonitorActionsEnabled(true);
            }
        }

        private async Task<bool> RunAgentTaskAsync(string ip, bool isMain, string mode)
        {
            string unit = $"{GetInspectionUnitFromIp(ip)} {ip} ({(isMain ? "Main" : "Vision")})";

            // 계정 리소스 키는 기존 패치 기능과 동일 규칙: SII는 Main/Vision 모두 INLINE.
            string accountType = mode == "SII" ? "INLINE" : (isMain ? "MAIN" : "VISION");
            var resources = new ResourceManager("AutoPatcher.Resource.UserInfo", typeof(MainWindow).Assembly);
            string id = resources.GetString($"{mode.ToUpper()}_{accountType}_ID");
            string pw = resources.GetString($"{mode.ToUpper()}_{accountType}_PW");
            if (string.IsNullOrEmpty(id))
            {
                Log($"[Monitor] {unit} 계정 정보 없음 ({mode.ToUpper()}_{accountType}_ID)", LogLevel.ERROR);
                return false;
            }

            string taskName = isMain ? _monConfig.TaskNameMain : _monConfig.TaskNameVision;

            bool reachable = await Task.Run(() =>
            {
                try
                {
                    using (var ping = new Ping()) return ping.Send(ip, 2000).Status == IPStatus.Success;
                }
                catch (PingException) { return false; }
            });
            if (!reachable)
            {
                Log($"[Monitor] {unit} 응답 없음 (ping)", LogLevel.ERROR);
                return false;
            }

            var result = await Task.Run(() => RunSchtasks(ip, taskName, id, pw));
            if (result.Item1)
                Log($"[Monitor] {unit} '{taskName}' 실행 요청 성공");
            else
                Log($"[Monitor] {unit} '{taskName}' 실행 실패: {result.Item2}", LogLevel.ERROR);
            return result.Item1;
        }

        private static Tuple<bool, string> RunSchtasks(string ip, string taskName, string id, string pw)
        {
            var psi = new ProcessStartInfo("schtasks.exe", $"/run /tn \"{taskName}\" /s {ip} /u {id} /p \"{pw}\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage),
                StandardErrorEncoding = Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage),
            };

            using (var process = Process.Start(psi))
            {
                Task<string> stdout = process.StandardOutput.ReadToEndAsync();
                Task<string> stderr = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit(60000))
                {
                    try { process.Kill(); } catch (InvalidOperationException) { }
                    return Tuple.Create(false, "시간 초과(60초)");
                }
                string output = (stdout.Result + " " + stderr.Result).Trim();
                return Tuple.Create(process.ExitCode == 0, output);
            }
        }

        #endregion
    }
}

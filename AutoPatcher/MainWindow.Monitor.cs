using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace AutoPatcher
{
    /// <summary>모니터링 탭: 중앙 상태파일 기반 버전 모니터링과 배포 등록.</summary>
    public partial class MainWindow
    {
        private MonitorConfig _monConfig;
        private DispatcherTimer _monitorTimer;
        private bool _monitorRefreshing;
        private bool _monitorRefreshPending;
        private bool _monitorBusy;
        private string _startupMessage = string.Empty;
        private string _refreshError = string.Empty;

        private void InitMonitor()
        {
            _monConfig = MonitorConfig.Load();
            MonitorGrid.ItemsSource = DataGridItems;
            lblMonCentral.Text = $"중앙: {_monConfig.CentralServer}";

            _monitorTimer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(_monConfig.RefreshMinutes) };
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

        private void ShowMonitorMessage() =>
            lblMonMessage.Text = string.Join("  /  ", new[] { _startupMessage, _refreshError }.Where(m => !string.IsNullOrEmpty(m)));

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
                    lblMonUpdated.Text = $"자동 갱신 {_monConfig.RefreshMinutes}분 · 마지막 갱신 {DateTime.Now:HH:mm:ss}";

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

            _monitorBusy = true;
            btnMonDeploy.IsEnabled = false;
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
                btnMonDeploy.IsEnabled = true;
            }
            await RefreshMonitorAsync();
        }

        #endregion
    }
}

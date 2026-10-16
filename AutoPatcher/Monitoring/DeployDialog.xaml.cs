using System;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace AutoPatcher
{
    public partial class DeployDialog : Window
    {
        private readonly string _centralServer;
        private readonly string _mainPcType;
        private readonly string _visionPcType;

        public string TargetPcType { get; private set; }
        public string CurrentVersion { get; private set; }
        public string NewVersion { get; private set; }
        public string SourceDir { get; private set; }

        public DeployDialog(string centralServer, string mainPcType, string visionPcType, string defaultSource)
        {
            InitializeComponent();
            _centralServer = centralServer;
            _mainPcType = mainPcType;
            _visionPcType = visionPcType;

            rdTargetMain.Content = mainPcType;
            rdTargetVision.Content = visionPcType;
            txtSource.Text = defaultSource?.TrimEnd('\\') ?? string.Empty;
            UpdateTarget();
        }

        private string SelectedPcType => rdTargetVision.IsChecked == true ? _visionPcType : _mainPcType;

        private string SelectedRoot => Path.Combine(_centralServer, SelectedPcType);

        private void Target_Checked(object sender, RoutedEventArgs e) => UpdateTarget();

        private void UpdateTarget()
        {
            if (txtTargetRoot == null || _centralServer == null) return;

            txtTargetRoot.Text = SelectedRoot;
            try
            {
                CurrentVersion = PackageRegistrar.ReadServerVersion(SelectedRoot);
                txtCurrentVersion.Text = string.IsNullOrEmpty(CurrentVersion) ? "(없음)" : CurrentVersion;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                CurrentVersion = string.Empty;
                txtCurrentVersion.Text = $"읽기 실패: {ex.Message}";
            }
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog { IsFolderPicker = true, Title = "패치 폴더 선택 (설비 bin에 덮어쓸 파일)" };
            if (Directory.Exists(txtSource.Text)) dialog.InitialDirectory = txtSource.Text;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok) txtSource.Text = dialog.FileName;
        }

        private void txtSource_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (txtSourceInfo == null) return;
            string dir = txtSource.Text.Trim();
            if (!Directory.Exists(dir))
            {
                txtSourceInfo.Text = string.IsNullOrEmpty(dir) ? string.Empty : "폴더가 없습니다.";
                return;
            }
            try
            {
                int count = Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories).Count();
                txtSourceInfo.Text = $"파일 {count}개 (updater.exe, .git 폴더는 제외하고 등록). 패키지는 설비 bin에 덮어쓰기만 하며 기존 파일을 지우지 않습니다.";
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                txtSourceInfo.Text = $"폴더 읽기 실패: {ex.Message}";
            }
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            string version = txtNewVersion.Text.Trim();
            string source = txtSource.Text.Trim().TrimEnd('\\');

            string error = Validate(version, source);
            if (error != null)
            {
                MessageBox.Show(error, "배포 등록", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            TargetPcType = SelectedPcType;
            NewVersion = version;
            SourceDir = source;
            DialogResult = true;
        }

        private string Validate(string version, string source)
        {
            if (string.IsNullOrEmpty(version)) return "새 버전을 입력하세요.";
            if (version.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return "버전에 폴더 이름으로 쓸 수 없는 문자가 있습니다.";
            if (string.Equals(version, CurrentVersion, StringComparison.OrdinalIgnoreCase)) return "현재 서버 버전과 같습니다.";
            if (Directory.Exists(PackageRegistrar.PackageDir(SelectedRoot, version))) return $"이미 존재하는 패키지입니다: package_{version}";
            if (!Directory.Exists(source)) return "패치 폴더가 없습니다.";
            if (!Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories).Any()) return "패치 폴더가 비어 있습니다.";
            return null;
        }
    }
}

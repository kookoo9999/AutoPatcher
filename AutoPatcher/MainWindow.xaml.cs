using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Management; // WMI 사용을 위한 네임스페이스
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Microsoft.Win32;
using System.IO.Compression;

namespace AutoPatcher
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool[] _ModeArray = new bool[4] {true,false,false,false};
        public bool[] ModeArray { get { return _ModeArray; } }
        public int SelectedMode { get { return Array.IndexOf(_ModeArray,true); } }

        public ObservableCollection<RowData> DataGridItems { get; private set; }
        public ObservableCollection<string> FileList { get; set; }

        private string _ExcelPath;
        public string ExcelPath
        {
            get { return _ExcelPath; }
            set { _ExcelPath = value; }
        }

        // IP 리스트 (로컬 네트워크 상의 PC들)
        static List<string> ipAddresses = new List<string>();

        // 배포할 프로그램 파일 경로
        static string sourceDirectory = @"D:\WPF\AutoPatch\AutoPatch\bin\Debug\net6.0-windows";
        
        // 백업 경로 (네트워크에서 접근 가능한 위치)
        static string backupRootPath = @"D:\WPF\TEST\Backup";
        static string processNameToCheck = "AutoPatch.exe";
        // 업데이트할 파일 리스트
        static List<string> filesToCheck;
        public MainWindow()
        {
            InitializeComponent();
            DataGridItems = new ObservableCollection<RowData>();
            FileList = new ObservableCollection<string>();
            //LoadExcelData(@"D:\WPF\test_lists.xlsx"); // 엑셀 파일 경로
            //SetupDataGridGrouping();
            DataGrid.ItemsSource = DataGridItems;            
            FileListBox.ItemsSource = FileList;            
        }

        public void Initialize()
        {
            DataGridItems.Clear();
            ipAddresses.Clear();
        }

        public  bool StartAutoPatch()
        {
            if (sourceDirectory == null)
            {
                MessageBox.Show("패치파일이 들어있는 폴더가 지정되지 않았습니다");
                return false;
            }
            
            if(ipAddresses.Count == 0)
            {
                MessageBox.Show("패치할 설비가 지정되지 않았습니다");
                return false;
            }
            
            List<string> filesToCheck = GetFileListFromDirectory(sourceDirectory);
            foreach (string ip in ipAddresses)
            {
                string remoteFolderPath = $@"\\{ip}\\main\\bin";
                string remoteBackupPath = remoteFolderPath + "../back_up\\";
                lblStatus.Content = $"access to [{ip}]...";
                Debug.WriteLine($"[{ip}] 접근 중...");

                try
                {
                    // 프로그램 실행 중 확인
                    if (IsProgramRunning(ip, processNameToCheck))
                    {
                        lblStatus.Content = $"[{ip}] program is running : Waiting for exit {processNameToCheck}...";
                        Debug.WriteLine($"[{ip}] 프로그램 실행 중: {processNameToCheck}. 종료를 기다립니다...");
                        WaitForProcessToExit(ip, processNameToCheck, timeoutSeconds: 120);
                    }
                    else
                    {
                        Debug.WriteLine($"[{ip}] 프로그램이 실행 중이 아닙니다. 패치 진행.");
                    }

                    BackupFolder(remoteFolderPath, remoteBackupPath);

                    // 파일 업데이트 작업
                    foreach (string fileName in filesToCheck)
                    {
                        string localFilePath = System.IO.Path.Combine(sourceDirectory, fileName);
                        string remoteFilePath = System.IO.Path.Combine(remoteFolderPath, fileName);

                        if (!File.Exists(remoteFilePath))
                        {
                            Debug.WriteLine($"[{ip}] 원격 파일 없음: {fileName}. 건너뜁니다.");
                            continue;
                        }

                        if (IsFileUpdateNeeded(localFilePath, remoteFilePath))
                        {
                            Debug.WriteLine($"[{ip}] 업데이트 필요: {fileName}. 백업 및 교체를 진행합니다.");
                            BackupAndReplaceFile(remoteFolderPath, fileName, localFilePath);
                        }
                        else
                        {
                            Debug.WriteLine($"[{ip}] 최신 상태 유지: {fileName}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[{ip}] 오류 발생: {ex.Message}");
                }
            }
            
            Debug.WriteLine("모든 작업이 완료되었습니다.");
            Console.ReadLine();

            return true;
        }

        public  string GetRelativePath(string basePath, string targetPath)
        {
            Uri baseUri = new Uri(AppendDirectorySeparator(basePath));
            Uri targetUri = new Uri(AppendDirectorySeparator(targetPath));

            Uri relativeUri = baseUri.MakeRelativeUri(targetUri);
            return Uri.UnescapeDataString(relativeUri.ToString().Replace('/', System.IO.Path.DirectorySeparatorChar));
        }

        private  string AppendDirectorySeparator(string path)
        {
            if (!path.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
            {
                path += System.IO.Path.DirectorySeparatorChar;
            }
            return path;
        }

        // 디렉터리의 모든 파일 리스트 가져오기
        List<string> GetFileListFromDirectory(string directoryPath)
        {
            try
            {
                // 디렉터리 내 모든 파일의 이름을 상대 경로로 반환
                List<string> files = new List<string>();
                foreach (string filePath in Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories)) //"*.dll
                {
                    // 파일 이름만 추가 (상대 경로)
                    string relativePath = GetRelativePath(directoryPath, filePath);
                    files.Add(relativePath);
                }
                return files;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"디렉터리 접근 중 오류 발생: {ex.Message}");
                return new List<string>();
            }
        }

        // 원격 프로그램 실행 여부 확인 (wmic로 수정필요)
        bool IsProgramRunning(string ip, string processName)
        {
            string command = $"Get-Process -Name {processName} -ComputerName {ip}";
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-Command \"{command}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();

                process.WaitForExit();
                if (!string.IsNullOrEmpty(error))
                {
                    Debug.WriteLine($"PowerShell 오류: {error}");
                    return false;
                }

                return !string.IsNullOrEmpty(output);
            }
            //try
            //{
            //    string query = $"SELECT * FROM Win32_Process WHERE Name LIKE '{processName}.exe'";
            //    ManagementScope scope = new ManagementScope($@"\\{ip}\root\cimv2");
            //    scope.Connect();

            //    ObjectQuery objQuery = new ObjectQuery(query);
            //    using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, objQuery))
            //    using (ManagementObjectCollection results = searcher.Get())
            //    {
            //        return results.Count > 0;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Debug.WriteLine($"[{ip}] WMI 오류 발생: {ex.Message}");
            //    return false;
            //}
        }

        // 원격 프로그램 종료 대기
        void WaitForProcessToExit(string ip, string processName, int timeoutSeconds)
        {
            int waitedSeconds = 0;

            while (IsProgramRunning(ip, processName))
            {
                if (waitedSeconds >= timeoutSeconds)
                {
                    Debug.WriteLine($"[{ip}] 타임아웃: {processName} 종료를 기다리는 동안 시간이 초과되었습니다.");
                    break;
                }

                Debug.WriteLine($"[{ip}] {processName} 실행 중... {waitedSeconds + 1}초 대기.");
                System.Threading.Thread.Sleep(1000);
                waitedSeconds++;
            }
        }

        // 파일 업데이트 필요 여부 확인
        bool IsFileUpdateNeeded(string localFilePath, string remoteFilePath)
        {
            // 로컬 파일 버전 가져오기
            string localVersion = GetFileVersion(localFilePath);
            string remoteVersion = GetFileVersion(remoteFilePath);

            if (!string.IsNullOrEmpty(localVersion) && !string.IsNullOrEmpty(remoteVersion))
            {
                // 버전 비교
                return IsNewerVersion(localVersion, remoteVersion);
            }
            else
            {
                // 버전 정보가 없으면 파일 수정 날짜로 비교
                DateTime localModified = File.GetLastWriteTime(localFilePath);
                DateTime remoteModified = File.GetLastWriteTime(remoteFilePath);

                return localModified > remoteModified;
            }
        }

        // 파일 버전 불러오기
        string GetFileVersion(string filePath)
        {
            if (File.Exists(filePath))
            {
                FileVersionInfo info = FileVersionInfo.GetVersionInfo(filePath);
                return info.FileVersion ?? string.Empty;
            }
            return string.Empty;
        }

        // 버전 비교 (true면 localVersion이 더 최신)
        static bool IsNewerVersion(string localVersion, string remoteVersion)
        {
            Version local = new Version(localVersion);
            Version remote = new Version(remoteVersion);
            return local.CompareTo(remote) > 0;
        }

        bool BackupFolder(string src, string des)
        {

            if (string.IsNullOrEmpty(src) || string.IsNullOrEmpty(des))
            {
                MessageBox.Show("Please provide both source and target directory paths.");
                return false;
            }

            try
            {
                CompressDirectory(src, des);
                MessageBox.Show("Directory compressed successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
                return false;
            }
            return true;
        }

        private void CompressDirectory(string sourceDir, string targetDir)
        {
            // 소스 디렉터리 경로가 올바른지 확인
            if (!Directory.Exists(sourceDir))
            {
                throw new DirectoryNotFoundException($"The source directory '{sourceDir}' does not exist.");
            }

            // 대상 디렉터리가 존재하지 않으면 생성
            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            // 압축된 파일 이름은 소스 디렉터리의 이름 + .zip 확장자를 사용
            string zipFilePath = System.IO.Path.Combine(targetDir, DateTime.Now.ToString("yyMMdd") + "_back.zip");

            if(File.Exists(zipFilePath))
            {
                zipFilePath = zipFilePath.Substring(0,zipFilePath.Length - 4);
                zipFilePath += "_" + DateTime.Now.ToString("HH_mm_ss") + ".zip";
            }
            // 디렉터리를 압축
            ZipFile.CreateFromDirectory(sourceDir, zipFilePath, CompressionLevel.Fastest, true);
        }

        // 백업 및 파일 교체
        static void BackupAndReplaceFile(string remoteFolderPath, string fileName, string localFilePath)
        {
            string remoteFilePath = System.IO.Path.Combine(remoteFolderPath, fileName);
            string backupPath = System.IO.Path.Combine(backupRootPath, DateTime.Now.ToString("yyMMdd")) + "_back";
            Directory.CreateDirectory(backupPath);

            string backupFilePath = System.IO.Path.Combine(backupPath, fileName);
            File.Move(remoteFilePath, backupFilePath);
            File.Copy(localFilePath, remoteFilePath);
            Debug.WriteLine($"[{remoteFolderPath}] 업데이트 완료: {fileName}");
        }

        private void LoadExcelData(string filePath)
        {
            DataGridItems.Clear();
            var excelApp = new Excel.Application();
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Open(filePath);                
                worksheet = workbook.Sheets[SelectedMode+1] as Excel.Worksheet;

                Excel.Range usedRange = worksheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;
                string temp_name = "";
                for (int row = 2; row <= rowCount; row++) // 1행은 헤더로 스킵
                {                    
                    string group_name = Convert.ToString(usedRange.Cells[row, 2]?.Value);
                    
                    if (string.IsNullOrEmpty(group_name)) { group_name = temp_name; }
                    else { temp_name = group_name; }
                    
                    DataGridItems.Add(new RowData
                    {
                        Group = group_name,
                        Server = Convert.ToString(usedRange.Cells[row, 1]?.Value),
                        LocalIP = Convert.ToString(usedRange.Cells[row, 3]?.Value),
                        InspectionUnit = Convert.ToString(usedRange.Cells[row, 4]?.Value),
                        PC1 = Convert.ToString(usedRange.Cells[row, 5]?.Value),
                        PC2 = Convert.ToString(usedRange.Cells[row, 6]?.Value),
                        PC3 = Convert.ToString(usedRange.Cells[row, 7]?.Value),
                        PC4 = Convert.ToString(usedRange.Cells[row, 8]?.Value)
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Excel data: {ex.Message}");
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();
            }
        }

        private void SetupDataGridGrouping()
        {
            //CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(DataGridItems);
            //PropertyGroupDescription groupDescription = new PropertyGroupDescription("Group");
            //view.GroupDescriptions.Add(groupDescription);

            // CollectionView를 가져와서 그룹화를 설정
            ICollectionView view = CollectionViewSource.GetDefaultView(DataGridItems);
            view.GroupDescriptions.Clear();
            view.GroupDescriptions.Add(new PropertyGroupDescription("Group"));
        }

        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            // Use FolderBrowserDialog to select a folder
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    LoadFilesFromFolder(dialog.SelectedPath);
                    lblCurDir.Content = dialog.SelectedPath;
                    sourceDirectory = dialog.SelectedPath;
                }
            }
        }

        private void btnLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedMode == -1)
            {
                MessageBox.Show("패치할 작업을 선택해주세요");
                return;
            }
            OpenFileDialog dlg = new OpenFileDialog();
            if(dlg.ShowDialog() ==true)
            {   
                if(dlg.CheckFileExists==true)
                {
                    ExcelPath = dlg.FileName;
                    LoadExcelData(dlg.FileName);
                    lblCurExcel.Content = dlg.FileName;                    
                }                
            }
        }

        private void LoadFilesFromFolder(string folderPath)
        {
            try
            {
                FileList.Clear(); // Clear previous files

                // Get all files in the folder
                var files = Directory.GetFiles(folderPath);

                foreach (var file in files)
                {
                    FileList.Add(file);                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading files: {ex.Message}");
            }
        }

        private void btnRunPatch_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedMode == -1)
            {
                MessageBox.Show("패치할 작업을 선택해주세요");
                return;
            }
            lblStatus.Content = "Patching..";
            if (StartAutoPatch()) lblStatus.Content = "Complete";
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as CheckBox;
            var dataItem = checkBox.DataContext as RowData;
            
            string ip = "";

            bool flag  = (checkBox.IsChecked == true) ? true : false;   
            if (checkBox.Name.Contains("Main") && dataItem != null)
            {
                ip = dataItem.PC1;
                processNameToCheck = "HDSInspector.exe";
            }

            else if (checkBox.Name.Contains("V1") && dataItem != null)
            {
                ip = dataItem.PC2;
                processNameToCheck = "IS.exe";
            }

            else if (checkBox.Name.Contains("V2") && dataItem != null)
            {
                ip = dataItem.PC3;
                processNameToCheck = "IS.exe";
            }

            else if (checkBox.Name.Contains("V3") && dataItem != null)
            {
                ip = dataItem.PC4;
                processNameToCheck = "IS.exe";
            }

            if (flag) ipAddresses.Add(ip);
            else ipAddresses.Remove(ip);
            
        }

        private void Radiobutton_Checked(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < 4; i++) _ModeArray[i] = false;
            var btn = sender as RadioButton;
            if (btn.Name.Contains("LF"))
            {
                _ModeArray[0] = true;
            }            
            else if (btn.Name.Contains("SII")) _ModeArray[1] = true;
            else if (btn.Name.Contains("BGA")) _ModeArray[2] = true;
            else if (btn.Name.Contains("COB")) _ModeArray[3] = true;

            if (!string.IsNullOrEmpty(ExcelPath))
            {
                Initialize();
                LoadExcelData(ExcelPath);
            }
        }

    }
    public class RowData : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool v1Selected;
        private bool v2Selected;
        private bool v3Selected;
        private bool mainSelected;

        public bool MainSelected 
        { 
            get { return mainSelected; }
            set { mainSelected = value; OnPropertyChanged(nameof(MainSelected)); }
        }
        public bool V1Selected 
        {
            get { return v1Selected; }
            set { v1Selected = value; OnPropertyChanged(nameof(V1Selected)); }
        }
        public bool V2Selected 
        {
            get { return v2Selected; }
            set { v2Selected = value; OnPropertyChanged(nameof(V2Selected)); }
        }
        public bool V3Selected 
        {
            get { return v3Selected; }
            set { v3Selected = value; OnPropertyChanged(nameof(V3Selected)); }
            
        }

        public string Group { get; set; }
        public string Server { get; set; }
        public string LocalIP { get; set; }
        public string InspectionUnit { get; set; }
        public string PC1 { get; set; }
        public string PC2 { get; set; }
        public string PC3 { get; set; }
        public string PC4 { get; set; }        

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));           
        }
    }
}

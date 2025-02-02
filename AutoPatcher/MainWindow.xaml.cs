﻿using System;
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
using System.Reflection;
using System.Resources;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using System.Security.Cryptography;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Collections;
using Microsoft.WindowsAPICodePack.Dialogs;
using static System.Net.WebRequestMethods;

namespace AutoPatcher
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        #region member

        #region radio button (LF,SII,BGA,COB)
        private bool[] _ModeArray = new bool[4] { false, false, false, false };
        public bool[] ModeArray { get { return _ModeArray; } }
        public int SelectedMode { get { return Array.IndexOf(_ModeArray, true); } }
        #endregion

        #region radio button (Main,Vision)
        private bool[] _TypeArray = new bool[2];
        public bool[] TypeArray { get { return _TypeArray; } }
        public int SelectType { get { return Array.IndexOf(_TypeArray, true); } }
        #endregion

        #region set all check (main,vision1,2,3)
        private bool[] _isAllSelected = new bool[4] { false, false, false, false };
        public bool[] IsAllSelected { get { return _isAllSelected; } }
        public int AllSelected { get { return Array.IndexOf(_isAllSelected, true); } }
        #endregion

        public ObservableCollection<RowData> DataGridItems { get; private set; }
        public ObservableCollection<string> FileList { get; set; }

        

        enum LogLevel
        {
            DEBUG = 0,
            INFO = 1,
            IMPORTANT = 2,
            ERROR = 3,
        }

        private string _ExcelPath;
        public string ExcelPath
        {
            get { return _ExcelPath; }
            set { _ExcelPath = value; }
        }

        private string _strType;
        public string PCType
        {
            get { return _strType; }
            set { _strType = value; }
        }

        private string _strMode;
        public string ModeType
        {
            get { return _strMode; }
            set { _strMode = value; }
        }

        // IP 리스트 (로컬 네트워크 상의 PC들)
        private List<string> _ipAddresses = new List<string>();
        public List<string> IPAddresses
        {
            get { return _ipAddresses; }
            set { _ipAddresses = value; }
        }

        // 배포할 프로그램 파일 경로
        private string _sourceDirectory;
        public string SourceDirectory
        {
            get { return _sourceDirectory; }
            set { _sourceDirectory = value; }
        }

        // 백업 경로 (네트워크에서 접근 가능한 위치)

        private string _processNameToCheck;
        public string ProcessNameToCheck
        {
            get { return _processNameToCheck; }
            set { _processNameToCheck = value; }
        }

        // 업데이트할 파일 리스트
        private List<string> _filesToCheck;
        public List<string> FilesToCheck
        {
            get { return _filesToCheck; }
            set { _filesToCheck = value; }
        }

        private List<string> _foldersToCheck;
        public List<string> FoldersToCheck
        {
            get { return _foldersToCheck; }
            set { _foldersToCheck = value; }
        }

        double originWidth, originHeight;
        ScaleTransform scale = new ScaleTransform();

        #endregion

        public MainWindow()
        {
            InitializeComponent();
            DataGridItems           = new ObservableCollection<RowData>();
            FileList                = new ObservableCollection<string>();            
            FilesToCheck            = new List<string>();
            IPAddresses             = new List<string>();
            FoldersToCheck          = new List<string>();            
            DataGrid.ItemsSource    = DataGridItems;
            FileListBox.ItemsSource = FileList;            
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            originHeight = this.Height;
            originWidth = this.Width;
            if (this.WindowState == WindowState.Maximized)
            {
                ChangeSize(this.ActualWidth, this.ActualHeight);
            }
            this.SizeChanged += new SizeChangedEventHandler(MainWindow_SizeChanged);
        }

        void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            ChangeSize(e.NewSize.Width, e.NewSize.Height);
        }

        private void ChangeSize(double width, double height)
        {
            scale.ScaleX = width / originWidth;
            scale.ScaleY = height / originHeight;

            FrameworkElement rootElement = this.Content as FrameworkElement;
            rootElement.LayoutTransform = scale;
        }

        public void InitItems()
        {
            DataGridItems.Clear();
            IPAddresses.Clear();
        }

        #region Update result
        
        private void Log(string msg, LogLevel level = LogLevel.INFO)
        {
            string text = $"[{DateTime.Now.ToString("g")}] [{level}] : {msg}\n";

            Run run = new Run(text);
            
            //LogBox.AppendText(text);            
            if (level == LogLevel.INFO || level == LogLevel.DEBUG)
            {
                run.Foreground = Brushes.Black;                
                SetMessage(msg);
            }
            else if (level == LogLevel.ERROR || level == LogLevel.IMPORTANT)
            {
                if (level == LogLevel.IMPORTANT) run.Foreground = Brushes.Orange;
                else run.Foreground = Brushes.Red;
                SetWarnning(msg);
            }
            LogBox.Inlines.Add(run);
            LogScroll.ScrollToEnd();
            return;
        }

        public void SetMessage(string msg)
        {
            this.lblStatus.Content = msg;
            this.lblStatus.Foreground = new SolidColorBrush(Colors.Black);
        }

        public void SetWarnning(string msg)
        {
            this.lblStatus.Content = msg;
            this.lblStatus.Foreground = new SolidColorBrush(Colors.Red);
        }

        public void SetComplete(string ip)
        {
            ChangeCellColor(ip, Brushes.OrangeRed);
        }

        public void SetFail(string ip)
        {
            ChangeCellColor(ip, Brushes.LimeGreen);
        }

        private void ChangeCellColor(string val, SolidColorBrush color)
        {
            // DataGrid의 각 행과 셀을 순회합니다.
            foreach (var row in this.DataGrid.Items)
            {
                var dataGridRow = (DataGridRow)this.DataGrid.ItemContainerGenerator.ContainerFromItem(row);
                if (dataGridRow == null) continue;

                foreach (System.Windows.Controls.DataGridCell cell in GetVisualChildren<System.Windows.Controls.DataGridCell>(dataGridRow))
                {
                    var column = this.DataGrid.Columns[cell.Column.DisplayIndex];

                    // 셀의 내용을 가져옵니다.
                    var cellContent = column.GetCellContent(row);
                    if (cellContent is ContentPresenter cp)
                    {

                        var sp = cp.ContentTemplate.FindName("StackPanel", cp) as StackPanel;
                        if (sp != null)
                        {
                            var textBlock = sp.Children.OfType<TextBlock>().FirstOrDefault();
                            if (textBlock != null && textBlock.Text == val)
                            {
                                // 텍스트가 일치하면 셀 배경 색상 변경
                                cell.Background = color; // 원하는 색상
                                break;
                            }
                        }
                    }
                }
            }
        }

        private IEnumerable<T> GetVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                var child = VisualTreeHelper.GetChild(depObj, i);
                if (child is T)
                    yield return (T)child;

                foreach (var childOfChild in GetVisualChildren<T>(child))
                {
                    yield return childOfChild;
                }
            }
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
                worksheet = workbook.Sheets[SelectedMode + 1] as Excel.Worksheet;

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
                        Group          = group_name,                        
                        InspectionUnit = Convert.ToString(usedRange.Cells[row, 3]?.Value),
                        PC1            = Convert.ToString(usedRange.Cells[row, 4]?.Value),
                        PC2            = Convert.ToString(usedRange.Cells[row, 5]?.Value),
                        PC3            = Convert.ToString(usedRange.Cells[row, 6]?.Value),
                        PC4            = Convert.ToString(usedRange.Cells[row, 7]?.Value),
                        PC5            = Convert.ToString(usedRange.Cells[row, 8]?.Value),
                        PC6            = Convert.ToString(usedRange.Cells[row, 9]?.Value)
                    });
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error loading Excel data: {ex.Message}");
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

        #endregion

        #region Patch
        /// <summary>
        /// Start Patch in ip
        /// </summary>
        /// <param name="ip">ip</param>
        /// <param name="BackupPathList">(bin,config)</param>
        /// <param name="remoteBackupPath">(\\backup)</param>
        /// <param name="remoteFolderPath">Excute path (bin)</param>
        /// <returns></returns>
        private bool Patch(string ip ,string[] BackupPathList,string remoteBackupPath,string remoteFolderPath)
        {
            try
            {
                #region ping
                if (!GetPingResult(ip))
                {
                    SetWarnning($"[{ip}] _ No response ");
                    System.Windows.MessageBox.Show($"No response {ip}");
                    ChangeCellColor(ip, Brushes.IndianRed);
                    return false;
                }
                #endregion

                #region check process running
                // 프로그램 실행 중 확인
                if (IsProgramRunning(ip, ProcessNameToCheck))
                {                    
                    Log($"[{ip}] {ProcessNameToCheck} is running : Waiting for exit ...");
                    Debug.WriteLine($"[{ip}] 프로그램 실행 중: {ProcessNameToCheck}. 종료를 기다립니다...");
                    WaitForProcessToExit(ip, ProcessNameToCheck, timeoutSeconds: 120);
                }
                else
                {                    
                    Log($"[{ip}] : [{ProcessNameToCheck}] is not running. Try to patch..");
                    Debug.WriteLine($"[{ip}] 프로그램이 실행 중이 아닙니다. 패치 진행.");
                }
                #endregion

                #region backup
                foreach (string strBackUpPath in BackupPathList)
                {                                 
                    Log($"[{ip}] _ [{strBackUpPath}] back up..");
                    if (!CheckBackupFolder(strBackUpPath, remoteBackupPath))
                    {                        
                        Log($"{strBackUpPath} dosen't exist", LogLevel.ERROR);
                        return false;
                    }

                    else
                    {                        
                        Log($"[{ip}] _ [{strBackUpPath}] Backup successfully!");
                    }
                }
                #endregion

                #region Update
                // subfolder 없을시 생성
                foreach (string folder in FoldersToCheck)
                {
                    if (!Directory.Exists(remoteFolderPath + "\\" + folder))
                    {
                        Directory.CreateDirectory(remoteFolderPath + "\\" + folder);
                    }
                }
                // 파일 업데이트 작업
                foreach (string fileName in FilesToCheck)
                {                    
                    Log($"[{ip}] _ Patching {fileName} ..");
                    string localFilePath = System.IO.Path.Combine(SourceDirectory, fileName);
                    string remoteFilePath = System.IO.Path.Combine(remoteFolderPath, fileName);

                    if (!System.IO.File.Exists(remoteFilePath))
                    {
                        Debug.WriteLine($"[{ip}] 원격 파일 없음: {fileName}.");
                    }

                    if (IsFileUpdateNeeded(localFilePath, remoteFilePath))
                    {                        
                        Debug.WriteLine($"[{ip}] 업데이트 필요: {fileName}. 백업 및 교체를 진행합니다.");
                        Log($"[{ip}] _ Update: {fileName}. Backup and replace in progress.");
                        ReplaceFile(ip,remoteFolderPath, fileName, localFilePath);
                    }
                    else
                    {
                        Debug.WriteLine($"[{ip}] 최신 상태 유지: {fileName}");
                        Log($"[{ip}] _ Skip : {fileName}");
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {                
                Log(ex.Message,LogLevel.ERROR);
                return false;
            }
            
            return true;
        }

        public bool StartAutoPatch()
        {
            if (SourceDirectory == null)
            {
                System.Windows.MessageBox.Show("No selected source directory");
                Log("No selected source directory", LogLevel.IMPORTANT);
                return false;
            }

            if (FilesToCheck.Count == 0)
            {                
                var res = GetFileListFromDirectory(SourceDirectory);
                FilesToCheck = res.files;
                FoldersToCheck = res.folders;
            }

            foreach (string ip in IPAddresses)
            {
                string strPCtype = PCType; //"Main" vision
                string diskType = "D";
                string remotePath = "";
                string[] remotePathList =
                {
                      $@"\\{ip}\\{strPCtype.ToLower()}",                       // ip : main,vision
                      $@"\\{ip}\\{strPCtype}",                                 // ip : Main,Vision
                      $@"\\{ip}\\{diskType}\\{strPCtype.ToLower()}",           // ip : D : main,vision
                      $@"\\{ip}\\{diskType}\\{strPCtype}",                     // ip : D : Main,Vision
                      $@"\\{ip}\\{diskType.ToLower()}\\{strPCtype.ToLower()}", // ip : d : main,vision
                      $@"\\{ip}\\{diskType.ToLower()}\\{strPCtype}",           // ip : d : Main,Vision
                };

                #region Set path (d:main , D:main, d:Main, D:Main)
                try
                {
                    foreach (string path in remotePathList)
                    {
                        if (Directory.Exists(path))
                        {
                            remotePath = path;                            
                            Log($"[{ip}] _ set path : [{remotePath}] ");
                            break;
                        }
                    }
                    if(string.IsNullOrEmpty(remotePath))
                    {                        
                        Log($"[{ip}] _ Does not exist path",LogLevel.IMPORTANT);
                        System.Windows.MessageBox.Show($"[{ip}] _ Does not exist path");
                        continue;
                    }
                }
                catch
                {
                    Log($"[{ip}] _ Does not exist path", LogLevel.ERROR);
                    System.Windows.MessageBox.Show($"[{ip}] _ Does not exist path");                                        
                    continue;
                }
                #endregion

                #region set path(backup,source)
                string remoteFolderPath = remotePath + "\\bin";
                string remoteBackupPath = remotePath + "\\backup";
                string remoteConfigPath = remotePath + "\\config";
                string[] BackupPathList = { remoteFolderPath, remoteConfigPath };
                #endregion

                #region Start Patch                
                Log($"Access to [{ip}]...");
                Debug.WriteLine($"[{ip}] 접근 중...");
                
                try
                {
                    if(Patch(ip,BackupPathList,remoteBackupPath,remoteFolderPath))
                    {                        
                        Log($"[{ip}] _ patch complete");
                        SetComplete(ip);
                    }
                    else
                    {                        
                        Log($"[{ip}] _ patch failed",LogLevel.IMPORTANT);
                        SetFail(ip);
                    }                    
                }
                catch (Exception ex)
                {                    
                    Log($"[{ip}] _ Error occured : {ex.Message}",LogLevel.ERROR);
                    Debug.WriteLine($"[{ip}] 오류 발생: {ex.Message}");
                    ChangeCellColor(ip, Brushes.IndianRed);
                    continue;
                }
                #endregion
            }

            Debug.WriteLine("모든 작업이 완료되었습니다.");
            Console.ReadLine();
            return true;
        }

        // 원격 프로그램 실행 여부 확인 (wmic로 수정필요)
        bool IsProgramRunning(string ip, string processName)
        {
            try
            {
                string remoteName = @"\\" + ip + @"\root\cimv2";

                ConnectionOptions con = new ConnectionOptions();
                ResourceManager rscManager = new ResourceManager("AutoPatcher.Resource.UserInfo", typeof(MainWindow).Assembly);

                con.Username = rscManager.GetString($"{ModeType.ToUpper()}_{PCType.ToUpper()}_ID");
                con.Password = rscManager.GetString($"{ModeType.ToUpper()}_{PCType.ToUpper()}_PW");

                ManagementScope managementScope = new ManagementScope(remoteName, con);
                managementScope.Options.Authentication = AuthenticationLevel.PacketPrivacy;
                managementScope.Connect();
                ObjectQuery objectQuery = new ObjectQuery($"SELECT * FROM Win32_Process Where Name = '{processName}'");
                ManagementObjectSearcher managementObjectSearcher = new ManagementObjectSearcher(managementScope, objectQuery);
                ManagementObjectCollection managementObjectCollection = managementObjectSearcher.Get();
                if (managementObjectCollection.Count > 0) return true;
                else return false;
            }
            catch
            {                
                Log("Check process error",LogLevel.ERROR);
                return false;
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
                    Log($"[{ip}] TimeOut: {processName} timeout occurred while waiting for termination.",LogLevel.ERROR);
                    Debug.WriteLine($"[{ip}] 타임아웃: {processName} 종료를 기다리는 동안 시간이 초과되었습니다.");
                    break;
                }                
                Log($"[{ip}] {processName} is running... waiting for {waitedSeconds + 1} sec..");
                Debug.WriteLine($"[{ip}] {processName} 실행 중... {waitedSeconds + 1}초 대기.");
                System.Threading.Thread.Sleep(1000);
                waitedSeconds++;
            }
        }

        // 파일 업데이트 필요 여부 확인
        bool IsFileUpdateNeeded(string localFilePath, string remoteFilePath)
        {
            // 로컬 파일 버전 가져오기
            //string localVersion = GetFileVersion(localFilePath);
            //string remoteVersion = GetFileVersion(remoteFilePath);

            //if (!string.IsNullOrEmpty(localVersion) && !string.IsNullOrEmpty(remoteVersion))
            //{
            //    // 버전 비교
            //    return IsNewerVersion(localVersion, remoteVersion);
            //}
            //else
            {
                // 버전 정보가 없으면 파일 수정 날짜로 비교
                DateTime localModified = System.IO.File.GetLastWriteTime(localFilePath);
                DateTime remoteModified = System.IO.File.GetLastWriteTime(remoteFilePath);

                return localModified > remoteModified;
            }
        }

        // 백업 및 파일 교체
        public void ReplaceFile(string ip, string remoteFolderPath, string fileName, string localFilePath)
        {
            string remoteFilePath = System.IO.Path.Combine(remoteFolderPath, fileName);
            //string backupPath = System.IO.Path.Combine(backupRootPath, DateTime.Now.ToString("yyMMdd")) + "_back";
            //Directory.CreateDirectory(backupPath);

            //string backupFilePath = System.IO.Path.Combine(backupPath, fileName);
            //File.Move(remoteFilePath, backupFilePath);
            System.IO.File.Delete(remoteFilePath);
            System.IO.File.Copy(localFilePath, remoteFilePath);                        
            Log($"[{ip}] _ [{remoteFolderPath}] Update completed: {fileName}");
            Debug.WriteLine($"[{remoteFolderPath}] 업데이트 완료: {fileName}");
        }

        #endregion

        #region Path,Directory,File

        private bool GetPingResult(string desIP)
        {
            try
            {
                using (Ping ping = new Ping())
                {
                    PingReply reply = ping.Send(desIP, 5000); // 5000ms

                    return reply.Status == IPStatus.Success;
                }
            }
            catch (Exception ex)
            {
                Log(ex.Message, LogLevel.ERROR);
                Console.WriteLine($"Error: {ex.Message}");
                return false;
            }
        }

        public string GetRelativePath(string basePath, string targetPath)
        {
            FileAttributes fa = System.IO.File.GetAttributes(targetPath);
            Uri baseUri = null;
            Uri targetUri = null;

            if ((fa & FileAttributes.Directory) == FileAttributes.Directory)
            {
                baseUri = new Uri(AppendDirectorySeparator(basePath));
                targetUri = new Uri(AppendDirectorySeparator(targetPath));

                //return Uri.UnescapeDataString(relativeUri.ToString().Replace('/', System.IO.Path.DirectorySeparatorChar));
            }
            else
            {
                baseUri = new Uri(basePath);
                targetUri = new Uri(targetPath);

                //return Uri.UnescapeDataString(relativeUri.ToString());
            }

            Uri relativeUri = baseUri.MakeRelativeUri(targetUri);
            string relativePath = Uri.UnescapeDataString(relativeUri.ToString().Replace('/', System.IO.Path.DirectorySeparatorChar));
            //relativePath = targetPath.Substring((basePath.Length+2));

            //if ((fa & FileAttributes.Directory) != FileAttributes.Directory)
            //{
            //    // 상대 경로에서 디렉터리 경로를 제외하고 파일명만 리턴
            //    return relativePath.Substring(relativePath.LastIndexOf(System.IO.Path.DirectorySeparatorChar) + 1);
            //}

            return relativePath;

        }

        private string AppendDirectorySeparator(string path)
        {
            if (!path.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
            {
                path += System.IO.Path.DirectorySeparatorChar;
            }
            return path;
        }

        private void LoadFilesFromFolder(string folderPath)
        {
            try
            {
                // Clear previous files
                FileList.Clear();
                FilesToCheck.Clear();
                FoldersToCheck.Clear();

                // Get all files in the folder
                var files = Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories);
                var folders = Directory.GetDirectories(folderPath, "*", SearchOption.AllDirectories);

                foreach (var folder in folders)
                {
                    FileList.Add(folder);
                }
                foreach (var file in files)
                {
                    if (file.Contains("CAM"))
                    {
                        FilesToCheck.Remove(file);
                        continue;
                    }
                    FileList.Add(file);
                }

                var res = GetFileListFromDirectory(SourceDirectory);
                FilesToCheck = res.files;
                FoldersToCheck = res.folders;

                Log($"{FilesToCheck.Count} files loaded");

            }
            catch (Exception ex)
            {
                Log($"Error loading files: {ex.Message}", LogLevel.ERROR);
                System.Windows.MessageBox.Show($"Error loading files: {ex.Message}");
            }
        }

        // 디렉터리의 모든 파일 리스트 가져오기
        (List<string> files, List<string> folders) GetFileListFromDirectory(string directoryPath)
        {
            try
            {
                // 디렉터리 내 모든 파일의 이름을 상대 경로로 반환
                List<string> folders = new List<string>();
                List<string> files = new List<string>();
                foreach (string folder in Directory.GetDirectories(directoryPath, "*", SearchOption.AllDirectories))
                {
                    if (folder.Contains("CAM")) continue;
                    string relative = GetRelativePath(directoryPath, folder);
                    folders.Add(relative);
                }
                foreach (string filePath in Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories)) //"*.dll
                {
                    if (filePath.Contains("CAM")) continue;
                    // 파일 이름만 추가 (상대 경로)
                    string relativePath = GetRelativePath(directoryPath, filePath);
                    files.Add(relativePath);
                }
                return (files, folders);
            }
            catch (Exception ex)
            {
                Log($"Error occurred while accessing the directory: {ex.Message}",LogLevel.ERROR);
                System.Windows.MessageBox.Show($"디렉터리 접근 중 오류 발생: {ex.Message}");
                return (new List<string>(), new List<string>());
            }
        }

        #endregion

        #region Backup

        // 파일 버전 불러오기
        string GetFileVersion(string filePath)
        {
            if (System.IO.File.Exists(filePath))
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

        public bool CheckBackupFolder(string src, string des)
        {

            if (string.IsNullOrEmpty(src) || string.IsNullOrEmpty(des))
            {
                Log("Please provide both source and target directory paths.", LogLevel.ERROR);
                System.Windows.MessageBox.Show("Please provide both source and target directory paths.");
                return false;
            }

            try
            {
                if (!Directory.Exists(src))
                {
                    Log($"{src} doesn't exist",LogLevel.ERROR);
                    System.Windows.MessageBox.Show($"{src} doesn't exist");
                    return false;
                }
                //CompressDirectory(src, des);
                BackupDirectory(src, des);

                //System.Windows.MessageBox.Show("Directory compressed successfully!");

            }
            catch (Exception ex)
            {
                Log($"Error: {ex.Message}", LogLevel.ERROR);
                System.Windows.MessageBox.Show($"Error: {ex.Message}");                
                return false;
            }
            return true;
        }

        private void BackupDirectory(string src, string dest)
        {
            if (string.IsNullOrEmpty(src) || string.IsNullOrEmpty(dest)) return;
            if (FilesToCheck.Count == 0) return;

            // bin,config 구분
            string backupType = (src.Contains("bin")) ? "bin" : "config";

            // 오늘 날짜 폴더 생성
            string backupPath = System.IO.Path.Combine(dest, DateTime.Now.ToString("yyMMdd"));
            if (!Directory.Exists(backupPath)) Directory.CreateDirectory(backupPath);

            // bin,config 생성
            backupPath = System.IO.Path.Combine(backupPath, backupType);
            if (!Directory.Exists(backupPath)) Directory.CreateDirectory(backupPath);

            // 폴더복사
            CopyFolder(src, backupPath);
        }

        public static void CopyFile(string src, string dest)
        {
            if (System.IO.File.Exists(src))
            {
                // 파일 생성
                System.IO.File.Copy(src, dest, true);
                Debug.WriteLine("File copied successfully!");
            }
            else
            {
                Debug.WriteLine("File not found at path: " + src);
            }
        }

        private void CopyFolder(string src, string dest)
        {
            if (Directory.Exists(src))
            {
                // 폴더 생성
                if (!Directory.Exists(dest)) Directory.CreateDirectory(dest);
                var files = Directory.GetFiles(src);

                // 파일 복사
                foreach (var file in files)
                {
                    string fileName = System.IO.Path.GetFileName(file);
                    string destinationFilePath = System.IO.Path.Combine(dest, fileName);
                    CopyFile(file, destinationFilePath);
                }

                // 폴더 복사 (재귀)
                var directories = Directory.GetDirectories(src);
                foreach (var directory in directories)
                {
                    string folderName = System.IO.Path.GetFileName(directory);
                    string destinationSubFolderPath = System.IO.Path.Combine(dest, folderName);
                    CopyFolder(directory, destinationSubFolderPath);
                }

                Debug.WriteLine("Folder copied successfully!");
            }
            else
            {
                Debug.WriteLine("Folder not found at path: " + src);
            }
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

            // bin,config 구분
            string backupType = (sourceDir.Contains("bin")) ? "_bin" : "_config";

            // 압축된 파일 이름은 소스 디렉터리의 이름 + .zip 
            string zipFilePath = System.IO.Path.Combine(targetDir, DateTime.Now.ToString("yyMMdd") + backupType + "_back.zip");

            if (System.IO.File.Exists(zipFilePath))
            {
                zipFilePath = zipFilePath.Substring(0, zipFilePath.Length - 4);
                zipFilePath += "_" + DateTime.Now.ToString("HH_mm_ss") + ".zip";
            }
            // 디렉터리를 압축
            ZipFile.CreateFromDirectory(sourceDir, zipFilePath, CompressionLevel.Fastest, true);
        }

        #endregion

        private void btnSetPatchDirectory(object sender, RoutedEventArgs e)
        {
            // Use FolderBrowserDialog to select a folder
            if (string.IsNullOrEmpty(ProcessNameToCheck))
            {
                Log("Set the process name",LogLevel.ERROR);                
                System.Windows.MessageBox.Show("Set the process name");
                return;
            }

            CommonOpenFileDialog cofd = new CommonOpenFileDialog();

            cofd.IsFolderPicker = true;

            if(cofd.ShowDialog()==CommonFileDialogResult.Ok)
            {
                SourceDirectory = cofd.FileName + "\\";
                Log($"Loading.. {SourceDirectory}");
                LoadFilesFromFolder(cofd.FileName);
                lblCurDir.Content = cofd.FileName;
                Log("Loaded patch list");                
            }
            

            //using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            //{
            //    System.Windows.Forms.DialogResult result = dialog.ShowDialog();

            //    if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
            //    {
            //        SourceDirectory = dialog.SelectedPath + "\\";
            //        LoadFilesFromFolder(dialog.SelectedPath);
            //        lblCurDir.Content = dialog.SelectedPath;
            //    }
                
            //}
            
        }

        private void btnLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedMode == -1)
            {
                Log("Select node type",LogLevel.ERROR);
                System.Windows.MessageBox.Show("Select node type");
                return;
            }

            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            if (dlg.ShowDialog() == true)
            {
                if (dlg.CheckFileExists == true)
                {
                    Log($"Loading...{dlg.FileName}");                                
                    ExcelPath = dlg.FileName;                    
                    LoadExcelData(dlg.FileName);
                    lblCurExcel.Content = dlg.FileName;
                    Log($"Loaded machine list : {ModeType}");                                  
                }
            }
        }

        private void btnRunPatch_Click(object sender, RoutedEventArgs e)
        {            
            if (SelectedMode == -1)
            {
                Log("Select a patch node", LogLevel.ERROR);
                System.Windows.MessageBox.Show("Select a patch node");                
                return;
            }

            var tempres = FilesToCheck.Find(x => x.Contains(ProcessNameToCheck));
            if (string.IsNullOrEmpty(tempres))
            {
                Log($"Does not including process : {ProcessNameToCheck}",LogLevel.ERROR);                
                System.Windows.MessageBox.Show($"Does not including process : {ProcessNameToCheck}");
                return;
            }

            if (IPAddresses.Count == 0)
            {
                System.Windows.MessageBox.Show("No selected mahcine");
                Log("No equipment specified", LogLevel.INFO);                
            }

            Log($"Start Patching .. {IPAddresses.Count} machines selected");            
            if (StartAutoPatch())
            {
                Log("All Patch completed");                
            }
            else
            {
                Log("Failed to patch",LogLevel.IMPORTANT);                
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as System.Windows.Controls.CheckBox;
            var dataItem = checkBox.DataContext as RowData;
            
            string ip = "";

            bool flag = (checkBox.IsChecked == true) ? true : false;
            if (checkBox.Name.Contains("Main") && dataItem != null)
            {
                ip = dataItem.PC1;
            }

            else if (checkBox.Name.Contains("V1") && dataItem != null)
            {
                ip = dataItem.PC2;
            }

            else if (checkBox.Name.Contains("V2") && dataItem != null)
            {
                ip = dataItem.PC3;
            }

            else if (checkBox.Name.Contains("V3") && dataItem != null)
            {
                ip = dataItem.PC4;
            }

            else if (checkBox.Name.Contains("V4") && dataItem != null)
            {
                ip = dataItem.PC5;
            }

            else if (checkBox.Name.Contains("V5") && dataItem != null)
            {
                ip = dataItem.PC6;
            }

            if (!string.IsNullOrEmpty(ip))
            {
                if (flag) IPAddresses.Add(ip);
                else IPAddresses.Remove(ip);
            }
        }

        private void AllCheckbox_checked(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as System.Windows.Controls.CheckBox;
            bool flag = (checkBox.IsChecked == true) ? true : false;
            if (checkBox.Name.Contains("Main"))
            {
                foreach (var item in DataGridItems)
                {
                    item.MainSelected = flag;
                }
            }
            else if (checkBox.Name.Contains("V1"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V1Selected = flag;
                }
            }
            else if (checkBox.Name.Contains("V2"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V2Selected = flag;

                }
            }
            else if (checkBox.Name.Contains("V3"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V3Selected = flag;
                }
            }

            else if (checkBox.Name.Contains("V4"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V4Selected = flag;
                }
            }

            else if (checkBox.Name.Contains("V5"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V5Selected = flag;
                }
            }

        }

        private void AllCheckbox_unchecked(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as System.Windows.Controls.CheckBox;

            if (checkBox.Name.Contains("Main"))
            {
                foreach (var item in DataGridItems)
                {
                    item.MainSelected = false;
                    IPAddresses.Add(item.PC1);
                }
            }
            else if (checkBox.Name.Contains("V1"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V1Selected = true;
                    IPAddresses.Add(item.PC2);
                }
            }
            else if (checkBox.Name.Contains("V2"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V2Selected = true;
                    IPAddresses.Add(item.PC3);
                }
            }
            else if (checkBox.Name.Contains("V3"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V3Selected = true;
                    IPAddresses.Add(item.PC4);
                }
            }

            else if (checkBox.Name.Contains("V4"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V4Selected = true;
                    IPAddresses.Add(item.PC5);
                }
            }

            else if (checkBox.Name.Contains("V5"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V5Selected = true;
                    IPAddresses.Add(item.PC6);
                }
            }
        }

        private void Radiobutton_Checked(object sender, RoutedEventArgs e)
        {
            var btn = sender as System.Windows.Controls.RadioButton;

            #region Datagrid Mode(LF,SII,BGA,COB)
            if (btn.GroupName.Contains("Mode"))
            {
                for (int i = 0; i < 4; i++) _ModeArray[i] = false;

                if (btn.Name.Contains("LF"))
                {
                    ModeType = "LF";
                    _ModeArray[0] = true;
                }
                else if (btn.Name.Contains("SII"))
                {
                    ModeType = "SII";
                    _ModeArray[1] = true;
                }

                else if (btn.Name.Contains("BGA"))
                {
                    ModeType = "BGA";
                    _ModeArray[2] = true;
                }

                else if (btn.Name.Contains("COB"))
                {
                    ModeType = "COB";
                    _ModeArray[3] = true;
                }

                if (!string.IsNullOrEmpty(ExcelPath))
                {                    
                    InitItems();
                    Log($"Loading...{ExcelPath}");
                    LoadExcelData(ExcelPath);                    
                    Log($"Loaded machine list : {ModeType}");
                }
            }
            #endregion
            #region PC Type(Main,Vision)
            else if (btn.GroupName.Contains("PCType"))
            {
                for (int i = 0; i < 2; i++) _TypeArray[i] = false;

                if (btn.Name.Contains("Main"))
                {
                    _TypeArray[0] = true;
                    PCType = "Main";
                    ProcessNameToCheck = "HDSInspector.exe";
                    lblProcName.Content = ProcessNameToCheck;
                }

                else if (btn.Name.Contains("Vision"))
                {
                    _TypeArray[1] = false;
                    PCType = "Vision";
                    ProcessNameToCheck = "IS.exe";
                    lblProcName.Content = ProcessNameToCheck;
                }
                Log($"{PCType}_{ProcessNameToCheck} selected ");
            }
            #endregion
        }

        private void chkMainAll_Checked(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as System.Windows.Controls.CheckBox;

            if (checkBox.Name.Contains("Main"))
            {
                foreach (var item in DataGridItems)
                {
                    item.MainSelected = false;
                    IPAddresses.Add(item.PC1);
                }
            }
            else if (checkBox.Name.Contains("V1"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V1Selected = true;
                    IPAddresses.Add(item.PC2);
                }
            }
            else if (checkBox.Name.Contains("V2"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V2Selected = true;
                    IPAddresses.Add(item.PC3);
                }
            }
            else if (checkBox.Name.Contains("V3"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V3Selected = true;
                    IPAddresses.Add(item.PC4);
                }
            }

            else if (checkBox.Name.Contains("V4"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V4Selected = true;
                    IPAddresses.Add(item.PC5);
                }
            }

            else if (checkBox.Name.Contains("V5"))
            {
                foreach (var item in DataGridItems)
                {
                    item.V5Selected = true;
                    IPAddresses.Add(item.PC6);
                }
            }
        }
    }
    public class RowData : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool v1Selected;
        private bool v2Selected;
        private bool v3Selected;
        private bool v4Selected;
        private bool v5Selected;
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

        public bool V4Selected
        {
            get { return v4Selected; }
            set { v4Selected = value; OnPropertyChanged(nameof(V4Selected)); }
        }

        public bool V5Selected
        {
            get { return v5Selected; }
            set { v5Selected = value; OnPropertyChanged(nameof(V5Selected)); }
        }

        public string Group { get; set; }
        public string InspectionUnit { get; set; }
        public string PC1 { get; set; }
        public string PC2 { get; set; }
        public string PC3 { get; set; }
        public string PC4 { get; set; }
        public string PC5 { get; set; }
        public string PC6 { get; set; }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

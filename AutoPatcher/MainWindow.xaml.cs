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
using System.Reflection;
using System.Resources;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using System.Security.Cryptography;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Collections;
using Microsoft.WindowsAPICodePack.Dialogs;
using static System.Net.WebRequestMethods;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Windows.Controls.Primitives;
using AutoPatcher.Properties;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Drawing;
using Common;

namespace AutoPatcher
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        #region member

        private bool _PatchResult;
        public bool PatchResult
        {
            get { return _PatchResult; }
            set { _PatchResult = value; }
        }

        private bool _IsPatching;
        public bool IsPatching
        {
            get { return _IsPatching; }
            set { _IsPatching = value; }
        }

        private bool _ErrorStop;
        public bool ErrorStop
        {
            get { return _ErrorStop; }
            set { _ErrorStop = value; }
        }

        
        double pgcnt = 0.0;
        double pgtotalcnt = 0.0;

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

        public Collection<CellData> CellDatas { get; set; }

        enum LogLevel
        {
            DEBUG = 0,
            INFO = 1,
            WARN = 2,
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
            CellDatas               = new Collection<CellData>();
            FileList                = new ObservableCollection<string>();            
            FilesToCheck            = new List<string>();
            IPAddresses             = new List<string>();
            FoldersToCheck          = new List<string>();            
            DataGrid.ItemsSource    = DataGridItems;
            FileListBox.ItemsSource = FileList;
        }

        string GetStringLogLevel(LogLevel lv)
        {
            string ret="";
            switch (lv)
            {
                case LogLevel.DEBUG:
                    ret = "DEBUG";
                    break;
                case LogLevel.INFO:
                    ret = " INFO ";
                    break;
                case LogLevel.WARN:
                    ret = " WARN ";
                    break;
                case LogLevel.ERROR:
                    ret = "ERROR";
                    break;
                default:
                    break;
            }
            return ret;
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
        
        private void Log_o(string msg, LogLevel level = LogLevel.INFO)
        {
            string text = $"[{DateTime.Now.ToString("g")}] [{GetStringLogLevel(level)}] : {msg}\n";

            Run run = new Run(text);
            
            //LogBox.AppendText(text);            
            if (level == LogLevel.INFO || level == LogLevel.DEBUG)
            {
                run.Foreground = System.Windows.Media.Brushes.AntiqueWhite;                
                //SetMessage(msg);
            }
            else if (level == LogLevel.ERROR || level == LogLevel.WARN)
            {
                if (level == LogLevel.WARN) run.Foreground = System.Windows.Media.Brushes.Orange;
                else run.Foreground = System.Windows.Media.Brushes.Red;
                //SetWarnning(msg);
            }
            LogBox.Inlines.Add(run);
            LogScroll.ScrollToEnd();

            return;
        }

        private void Log(string msg, LogLevel level = LogLevel.INFO)
        {
            string text = $"[{DateTime.Now.ToString("g")}] [{GetStringLogLevel(level)}] : {msg}\n"; //
            Run run = new Run(text); //

            if (level == LogLevel.INFO || level == LogLevel.DEBUG) //
            {
                run.Foreground = System.Windows.Media.Brushes.AntiqueWhite; //
            }
            else if (level == LogLevel.ERROR || level == LogLevel.WARN) //
            {
                if (level == LogLevel.WARN) run.Foreground = System.Windows.Media.Brushes.Orange; //
                else run.Foreground = System.Windows.Media.Brushes.Red; //
            }

            Dispatcher.InvokeAsync(() => // UI 스레드에서 실행 보장
            {
                LogBox.Inlines.Add(run); 
                LogScroll.ScrollToEnd(); 
            });
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
            IPAddresses.Remove(ip);
            CellData targetCell = CellDatas.FirstOrDefault(item => item.IP == ip);
            ChangeCellColor(targetCell.ROW, targetCell.COLUMN, System.Windows.Media.Brushes.DeepSkyBlue);            
            //ChangeCellColor(ip, Brushes.LimeGreen);
        }

        public void SetFail(string ip)
        {
            //IPAddresses.Remove(ip);
            CellData targetCell = CellDatas.FirstOrDefault(item => item.IP == ip);
            ChangeCellColor(targetCell.ROW, targetCell.COLUMN, System.Windows.Media.Brushes.IndianRed);
            //ChangeCellColor(ip, Brushes.IndianRed);
        }

        private void ChangeCellColor(string val, SolidColorBrush color)
        {            
            // DataGrid의 각 행과 셀을 순회            
            foreach (var row in this.DataGrid.Items)
            {
                var dataGridRow = (DataGridRow)this.DataGrid.ItemContainerGenerator.ContainerFromItem(row);
                if (dataGridRow == null) continue;

                foreach (System.Windows.Controls.DataGridCell cell in GetVisualChildren<System.Windows.Controls.DataGridCell>(dataGridRow))
                {
                    var column = this.DataGrid.Columns[cell.Column.DisplayIndex];

                    // 셀의 내용
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
                                cell.Focus();
                                break;
                            }
                        }
                    }
                }
            }


            
        }

        private void ChangeCellColor(int row, int col, SolidColorBrush color)
        {
            var cell = GetCell(row, col);
            if(cell != null)
            {
                cell.Background = color;
                cell.Focus();
            }
        }

        private System.Windows.Controls.DataGridCell GetCell(int rowIndex, int columnIndex)
        {
            var row = (DataGridRow)DataGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex);
            if (row == null) return null;

            var cell = DataGrid.Columns[columnIndex].GetCellContent(row).Parent as System.Windows.Controls.DataGridCell;
            return cell;
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
                    System.Action action = delegate
                    {
                        DataGridItems.Add(new RowData
                        {
                            Group = group_name,
                            InspectionUnit = Convert.ToString(usedRange.Cells[row, 3]?.Value),
                            PC1 = Convert.ToString(usedRange.Cells[row, 4]?.Value),
                            PC2 = Convert.ToString(usedRange.Cells[row, 5]?.Value),
                            PC3 = Convert.ToString(usedRange.Cells[row, 6]?.Value),
                            PC4 = Convert.ToString(usedRange.Cells[row, 7]?.Value),
                            PC5 = Convert.ToString(usedRange.Cells[row, 8]?.Value),
                            PC6 = Convert.ToString(usedRange.Cells[row, 9]?.Value)
                        });
                    };
                    this.Dispatcher.Invoke(action);
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

        public bool ConnectNetworkDrive(string IPPath, string ID, string Password)
        {

            int Count = 0;
            while (true)
            {
                Count++;
                int state = NetworkDriveConnector.TryConnectNetwork(IPPath, ID, Password);

                if (NetworkDriveConnector.TryConnectResult(state))
                {
                    return true;
                }
                else if (!NetworkDriveConnector.TryConnectResult(state) && Count >3)
                {
                    return false;
                }
            }
        }


        /// <summary>
        /// Start Patch in ip
        /// </summary>
        /// <param name="ip">ip</param>
        /// <param name="BackupPathList">(bin,config)</param>
        /// <param name="remoteBackupPath">(\\backup)</param>
        /// <param name="remoteFolderPath">Excute path (bin)</param>
        /// <returns></returns>
        private async Task Patch(string ip ,List<string> BackupPathList,string remoteBackupPath,string remoteFolderPath)
        {
            try
            {
                #region ping
                if (!GetPingResult(ip))
                {
                    SetWarnning($"[{ip}] _ No response ");
                    System.Windows.MessageBox.Show($"No response {ip}");
                    SetFail(ip);
                    //ChangeCellColor(ip, Brushes.IndianRed);
                    await Task.Delay(10);
                    PatchResult = false;
                    return ;
                }
                Log($"[{ip}] _ Ping success");
                await Task.Delay(10);
                #endregion

                #region check process running
                // 프로그램 실행 중 확인
                if (IsProgramRunning(ip, ProcessNameToCheck))
                {                    
                    Log($"[{ip}] _ {ProcessNameToCheck} is running : Waiting for exit ...");
                    Debug.WriteLine($"[{ip}] 프로그램 실행 중: {ProcessNameToCheck}. 종료를 기다립니다...");
                    await Task.Delay(10);
                    WaitForProcessToExit(ip, ProcessNameToCheck, timeoutSeconds: 120);
                }
                else
                {                    
                    Log($"[{ip}] _ [{ProcessNameToCheck}] is not running. Try to patch..");
                    await Task.Delay(10);
                    Debug.WriteLine($"[{ip}] 프로그램이 실행 중이 아닙니다. 패치 진행.");
                }
                #endregion

                #region backup
                foreach (string strBackUpPath in BackupPathList)
                {                                 
                    Log($"[{ip}] _ [{strBackUpPath}] back up..");
                    await Task.Delay(1);
                    if (!CheckBackupFolder(strBackUpPath, remoteBackupPath))
                    {                        
                        Log($"[{ip}] _ {strBackUpPath} dosen't exist", LogLevel.ERROR);
                        await Task.Delay(1);
                        PatchResult = false;
                        return ;
                    }

                    else
                    {                        
                        Log($"[{ip}] _ [{strBackUpPath}] Backup successfully!");
                        await Task.Delay(1);
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
                int cnt = 0; pgcnt = 0.0;
                foreach (string fileName in FilesToCheck)
                {                    
                    //Log($"[{ip}] _ Patching {fileName} ..");
                    await Task.Delay(10);
                    cnt++;
                    string localFilePath = System.IO.Path.Combine(SourceDirectory, fileName);
                    string remoteFilePath = System.IO.Path.Combine(remoteFolderPath, fileName);

                    //if (!System.IO.File.Exists(remoteFilePath))
                    //{
                    //    Debug.WriteLine($"[{ip}] 원격 파일 없음: {fileName}.");
                    //}

                    if (IsFileUpdateNeeded(localFilePath, remoteFilePath))
                    {                        
                        Debug.WriteLine($"[{ip}] 업데이트 필요: {fileName}. 백업 및 교체를 진행합니다.");
                        //Log($"[{ip}] _ Update: {fileName}. Backup and replace in progress.");
                        await Task.Delay(1);
                        ReplaceFile(ip,remoteFolderPath, fileName, localFilePath);
                    }
                    else
                    {
                        Debug.WriteLine($"[{ip}] 최신 상태 유지: {fileName}");
                        Log($"[{ip}] _ Skip : {fileName}");
                        await Task.Delay(1);
                    }
                    
                    pgcnt = (((double)cnt / (double)FilesToCheck.Count)) * 100;
                    pgtotalcnt += (((double)1 / ((double)FilesToCheck.Count * (double)IPAddresses.Count ))) * 100;

                    pbstatusBar.Value = pgcnt;
                    txtStatusBar.Text = pgcnt.ToString("0.0")+"%";
                    pbtotalBar.Value = pgtotalcnt;
                    txttotalBar.Text = pgtotalcnt.ToString("0.0") + "%";
                    await Task.Delay(1);
                }
                #endregion
            }
            catch (Exception ex)
            {                
                Log(ex.Message,LogLevel.ERROR);
                await Task.Delay(10);
                PatchResult = false;
                return ;
            }

            PatchResult = true;
            return ;
        }

        public async Task StartAutoPatch()
        {
            if (SourceDirectory == null)
            {
                Log("No selected source directory", LogLevel.WARN);
                System.Windows.MessageBox.Show("No selected source directory");
                return;
            }

            if (FilesToCheck.Count == 0)
            {                
                var res = GetFileListFromDirectory(SourceDirectory);
                FilesToCheck = res.files;
                FoldersToCheck = res.folders;
            }

            // UI 요소 초기화
            await Dispatcher.InvokeAsync(() => {
                //pbstatusBar.Visibility = Visibility.Visible; //
                txtStatusBar.Visibility = Visibility.Visible; //
                pbtotalBar.Visibility = Visibility.Visible; //
                txttotalBar.Visibility = Visibility.Visible; //
                pbtotalBar.Value = 0; //
                txttotalBar.Text = "0.0%"; //
            });
            pgtotalcnt = 0.0; //

            List<Task> patchTasks = new List<Task>();
            int totalIPs = IPAddresses.Count;
            int completedIPs = 0;

            // 동시 실행 작업 수 제한 (선택 사항)
            // int maxConcurrentTasks = 5; // 예: 동시에 5개 IP만 처리
            // SemaphoreSlim semaphore = new SemaphoreSlim(maxConcurrentTasks);

            #region 이전방식
            //foreach (string ip in IPAddresses)
            //{
            //    Log($"[{ip}] _ Try to access...");
            //    lblIP.Content = ip;
            //    pbstatusBar.Value = 0;
            //    txtStatusBar.Text = 0.ToString("0.0") + "%";

            //    CellData cellData = CellDatas.FirstOrDefault(item => item.IP == ip);
            //    var cell = GetCell(cellData.ROW, cellData.COLUMN);
            //    ChangeCellColor(cellData.ROW, cellData.COLUMN, System.Windows.Media.Brushes.LimeGreen);
            //    cell.Focus();
            //    await Task.Delay(100);

            //    #region ping
            //    if (!GetPingResult(ip))
            //    {
            //        //SetWarnning($"[{ip}] _ No response ");
            //        Log($"[{ip}] _ No response",LogLevel.ERROR);
            //        //System.Windows.MessageBox.Show($"No response {ip}");
            //        SetFail(ip);                    
            //        await Task.Delay(10);
            //        PatchResult = false;
            //        pgtotalcnt += ((double)1 / (double)(IPAddresses.Count))*100;
            //        pbtotalBar.Value = pgtotalcnt;
            //        txttotalBar.Text = pgtotalcnt.ToString("0.0") + "%";
            //        continue;
            //    }
            //    Log($"[{ip}] _ Ping success");
            //    await Task.Delay(1);
            //    #endregion

            //    string strPCtype = PCType; //"Main" vision
            //    if (ModeType == "SII" && ProcessNameToCheck == "IS.exe") strPCtype += "\\IS";
            //    string diskType = "D";
            //    string remotePath = "";
            //    string[] remotePathList =
            //    {
            //          $"\\\\{ip}\\{strPCtype.ToLower()}",                       // ip : main,vision
            //          $"\\\\{ip}\\{strPCtype}",                                 // ip : Main,Vision
            //          $"\\\\{ip}\\{diskType}\\{strPCtype.ToLower()}",           // ip : D : main,vision
            //          $"\\\\{ip}\\{diskType}\\{strPCtype}",                     // ip : D : Main,Vision
            //          $"\\\\{ip}\\{diskType.ToLower()}\\{strPCtype.ToLower()}", // ip : d : main,vision
            //          $"\\\\{ip}\\{diskType.ToLower()}\\{strPCtype}",           // ip : d : Main,Vision 
            //    };
            //    #region Set path (d:main , D:main, d:Main, D:Main)

            //    ResourceManager rscManager = new ResourceManager("AutoPatcher.Resource.UserInfo", typeof(MainWindow).Assembly);
            //    string id = rscManager.GetString($"{ModeType.ToUpper()}_{PCType.ToUpper()}_ID");
            //    string pw = rscManager.GetString($"{ModeType.ToUpper()}_{PCType.ToUpper()}_PW");


            //    try
            //    {
            //        foreach (string path in remotePathList)
            //        {
            //            #region check permission     
            //            if (!ConnectNetworkDrive(path, id, pw))
            //            {
            //                SetFail(ip);
            //                Log($"[{ip}] _ Failed to set permission");
            //                break;
            //            }
            //            #endregion

            //            #region check directory
            //            if (Directory.Exists(path))
            //            {
            //                remotePath = path;                            
            //                Log($"[{ip}] _ set path : [{remotePath}] ");
            //                await Task.Delay(10);
            //                break;
            //            }
            //            #endregion
            //        }
            //        if (string.IsNullOrEmpty(remotePath))
            //        {                        
            //            Log($"[{ip}] _ Does not exist path",LogLevel.WARN);
            //            System.Windows.MessageBox.Show($"[{ip}] _ Does not exist path");
            //            SetFail(ip);
            //            await Task.Delay(10);
            //            continue;
            //        }
            //    }
            //    catch
            //    {
            //        Log($"[{ip}] _ Does not exist path", LogLevel.ERROR);
            //        System.Windows.MessageBox.Show($"[{ip}] _ Does not exist path");
            //        await Task.Delay(1);
            //        continue;
            //    }
            //    #endregion

            //    #region set path(backup,source)
            //    string remoteFolderPath = remotePath + "\\bin";
            //    string remoteBackupPath = remotePath + "\\backup";
            //    string remoteConfigPath = remotePath + "\\config";                
            //    List<string> BackupPathList = new List<string>();

            //    BackupPathList.Add(remoteFolderPath);
            //    if (ProcessNameToCheck == "HDSInspector.exe") BackupPathList.Add(remoteConfigPath);

            //    #endregion

            //    #region Start Patch                

            //    Debug.WriteLine($"[{ip}] 접근 중...");

            //    try
            //    {
            //        await Patch(ip, BackupPathList, remoteBackupPath, remoteFolderPath);
            //        if(PatchResult)
            //        {                        
            //            Log($"[{ip}] _ patch complete");
            //            SetComplete(ip);
            //            if (!ProcessStart(ip)) Log($"[{ip}] _  failed to run process", LogLevel.ERROR);
            //            else Log($"[{ip}] _  success to run process");
            //            await Task.Delay(50);
            //        }
            //        else
            //        {                        
            //            Log($"[{ip}] _ patch failed",LogLevel.WARN);
            //            SetFail(ip);
            //            await Task.Delay(50);
            //        }                    
            //        await Task.Delay(50);
            //    }
            //    catch (Exception ex)
            //    {                    
            //        Log($"[{ip}] _ Error occured : {ex.Message}",LogLevel.ERROR);
            //        Debug.WriteLine($"[{ip}] 오류 발생: {ex.Message}");
            //        SetFail(ip);
            //        //ChangeCellColor(ip, Brushes.IndianRed);
            //        await Task.Delay(10);
            //        continue;
            //    }
            //    PatchResult = false;
            //    #endregion
            //}
            #endregion

            foreach (string ip in IPAddresses)
            {
                // await semaphore.WaitAsync(); // 동시 작업 수 제한 시 사용

                patchTasks.Add(ProcessSingleIPAsync(ip, totalIPs, () =>
                {
                    // 전체 진행률 스레드 안전하게 업데이트
                    Interlocked.Increment(ref completedIPs);
                    double currentTotalProgress = ((double)completedIPs / totalIPs) * 100.0;
                    Dispatcher.InvokeAsync(() =>
                    {
                        pbtotalBar.Value = currentTotalProgress;
                        txttotalBar.Text = currentTotalProgress.ToString("0.0") + "%";
                    });
                    // semaphore.Release(); // 동시 작업 수 제한 시 사용
                }));
            }

            await Task.WhenAll(patchTasks);

            await Dispatcher.InvokeAsync(() => lblList.Content = $"Number of left machine : {IPAddresses.Count}");
            
            await Dispatcher.InvokeAsync(() => Log("All patching operations have been attempted."));
            Debug.WriteLine("모든 작업이 완료되었습니다."); 
        }

        private async Task ProcessSingleIPAsync(string ip, int totalIPCount, System.Action onIpCompletedCallback)
        {
            string currentIp = ip;
            bool ipPatchSuccess = false;

            try
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    Log($"[{currentIp}] _ Try to access..."); //
                                                              // lblIP.Content = currentIp; // 병렬 처리 시 의미가 모호할 수 있음
                                                              // pbstatusBar.Value = 0; // 개별 IP 진행률 표시 방법 재고 필요
                                                              // txtStatusBar.Text = "0.0%"; //

                    CellData cellData = CellDatas.FirstOrDefault(item => item.IP == currentIp); //
                    if (cellData != null)
                    {
                        ChangeCellColor(cellData.ROW, cellData.COLUMN, System.Windows.Media.Brushes.LimeGreen); //
                     // GetCell(cellData.ROW, cellData.COLUMN)?.Focus(); // 포커싱은 UI를 혼란스럽게 할 수 있음
                    }
                });
                await Task.Delay(100); //

                if (!GetPingResult(currentIp)) //
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        Log($"[{currentIp}] _ No response", LogLevel.ERROR); //
                        SetFail(currentIp); //
                    });
                    return;
                }
                await Dispatcher.InvokeAsync(() => Log($"[{currentIp}] _ Ping success")); //

                // 경로 설정 및 연결 로직 (기존과 유사, 단 ConnectNetworkDrive는 스레드 안전하게 수정되었다고 가정)
                // ... (ConnectNetworkDrive, Directory.Exists 등 호출 부분)
                // PCType, ModeType 등 멤버 변수 사용 시, 병렬 작업 중 변경되지 않는다는 가정하에 사용
                string strPCtype = PCType; // main, vision
                if (ModeType == "SII" && ProcessNameToCheck == "IS.exe") strPCtype += "\\IS"; //
                string diskType = "D"; //
                string remotePath = ""; //
               
                string[] remotePathList =
                {
                      $"\\\\{ip}\\{strPCtype.ToLower()}",                       // ip : main,vision
                      $"\\\\{ip}\\{strPCtype}",                                 // ip : Main,Vision
                      $"\\\\{ip}\\{diskType}\\{strPCtype.ToLower()}",           // ip : D : main,vision
                      $"\\\\{ip}\\{diskType}\\{strPCtype}",                     // ip : D : Main,Vision
                      $"\\\\{ip}\\{diskType.ToLower()}\\{strPCtype.ToLower()}", // ip : d : main,vision
                      $"\\\\{ip}\\{diskType.ToLower()}\\{strPCtype}",           // ip : d : Main,Vision
                };


                ResourceManager rscManager = new ResourceManager("AutoPatcher.Resource.UserInfo", typeof(MainWindow).Assembly); //
                string id = rscManager.GetString($"{ModeType.ToUpper()}_{PCType.ToUpper()}_ID"); //
                string pw = rscManager.GetString($"{ModeType.ToUpper()}_{PCType.ToUpper()}_PW"); //

                bool pathFound = false;

                foreach (string pathAttempt in remotePathList) // remotePathList는 기존처럼 생성
                {
                    // 1. 네트워크 드라이브 연결 시도
                    bool connectionApiSuccess = ConnectNetworkDrive(pathAttempt, id, pw); //

                    if (!connectionApiSuccess)
                    {
                        // WNetUseConnection API 호출이 명시적으로 성공하지 않았음을 로그로 남깁니다.
                        // (예: 이미 연결됨, 다른 오류 등). 그러나 경로는 여전히 접근 가능할 수 있습니다.
                        await Dispatcher.InvokeAsync(() => Log($"[{currentIp}] _ ConnectNetworkDrive API for '{pathAttempt}' reported an issue or no new connection was made. Proceeding to check path accessibility.", LogLevel.INFO));
                    }
                    else
                    {
                        await Dispatcher.InvokeAsync(() => Log($"[{currentIp}] _ ConnectNetworkDrive API for '{pathAttempt}' reported success."));
                    }

                    // 2. 경로 존재 유무 및 접근 가능성 확인 
                    // ConnectNetworkDrive의 반환 값과 관계없이 실제 경로 접근을 시도
                    try
                    {
                        if (Directory.Exists(pathAttempt))
                        {
                            remotePath = pathAttempt; // 접근 가능한 경로를 찾음
                            await Dispatcher.InvokeAsync(() => Log($"[{currentIp}] _ Successfully accessed and set remote path: [{remotePath}]"));
                            pathFound = true;
                            break; // 유효한 경로를 찾았으므로 루프 종료
                        }
                        else
                        {
                            // 해당 경로가 존재하지 않거나 접근할 수 없음을 로그로 남깁니다.
                            await Dispatcher.InvokeAsync(() => Log($"[{currentIp}] _ Path '{pathAttempt}' not found or not accessible.", LogLevel.INFO));
                        }
                    }
                    catch (Exception ex)
                    {
                        // Directory.Exists 중 예외 발생 시 (네트워크 오류 등)
                        await Dispatcher.InvokeAsync(() => Log($"[{currentIp}] _ Exception while checking Directory.Exists for '{pathAttempt}': {ex.Message}", LogLevel.WARN));
                    }
                }

                if (!pathFound || string.IsNullOrEmpty(remotePath))
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        Log($"[{currentIp}] _ CRITICAL: No accessible remote path found after trying all candidates. Patching for this IP will be skipped.", LogLevel.ERROR);
                        SetFail(currentIp); //
                    });
                    // onIpCompletedCallback는 ProcessSingleIPAsync의 finally 블록에서 호출되므로 여기서 return해도 문제없음
                    return;
                }

                string remoteFolderPath = System.IO.Path.Combine(remotePath, "bin"); 
                string remoteBackupPath = System.IO.Path.Combine(remotePath, "backup"); 
                string remoteConfigPath = System.IO.Path.Combine(remotePath, "config");
                List<string> backupPathList = new List<string> { remoteFolderPath }; 
                if (ProcessNameToCheck == "HDSInspector.exe") backupPathList.Add(remoteConfigPath); 


                // Patch 메서드를 호출 (PatchInternalAsync로 리팩토링 버전)
                ipPatchSuccess = await PatchInternalAsync(currentIp, backupPathList, remoteBackupPath, remoteFolderPath, FilesToCheck, SourceDirectory, ProcessNameToCheck);

                if (ipPatchSuccess)
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        Log($"[{currentIp}] _ patch complete"); //
                        SetComplete(currentIp); //
                    });
                    if (!ProcessStart(currentIp)) //
                        await Dispatcher.InvokeAsync(() => Log($"[{currentIp}] _  failed to run process", LogLevel.ERROR)); //
                    else
                        await Dispatcher.InvokeAsync(() => Log($"[{currentIp}] _  success to run process")); //
                }
                else
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        Log($"[{currentIp}] _ patch failed", LogLevel.WARN); //
                        SetFail(currentIp); //
                    });
                }

                #region 체크박스 해제 로직 추가
                CellData cellInfoForUncheck = CellDatas.FirstOrDefault(item => item.IP == currentIp);
                if (cellInfoForUncheck != null)
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        // DataGridItems 컬렉션과 RowData 인덱스의 유효성 확인
                        if (cellInfoForUncheck.ROW >= 0 && cellInfoForUncheck.ROW < DataGridItems.Count)
                        {
                            RowData rowDataItem = DataGridItems[cellInfoForUncheck.ROW];

                            // RowData 객체 내에서 현재 IP와 일치하는 PC 속성을 찾아 해당 Selected 플래그를 false로 설정
                            if (rowDataItem.PC1 == currentIp)
                            {
                                rowDataItem.MainSelected = false;
                            }
                            else if (rowDataItem.PC2 == currentIp)
                            {
                                rowDataItem.V1Selected = false;
                            }
                            else if (rowDataItem.PC3 == currentIp)
                            {
                                rowDataItem.V2Selected = false;
                            }
                            else if (rowDataItem.PC4 == currentIp)
                            {
                                rowDataItem.V3Selected = false;
                            }
                            else if (rowDataItem.PC5 == currentIp)
                            {
                                rowDataItem.V4Selected = false;
                            }
                            else if (rowDataItem.PC6 == currentIp)
                            {
                                rowDataItem.V5Selected = false;
                            }
                            else
                            {
                                // 만약 RowData 내에서 해당 IP를 가진 PC 속성을 찾지 못한 경우 (데이터 불일치 등)
                                Log($"[{currentIp}] Could not find matching PC property in RowData to uncheck checkbox for IP {currentIp} in row {cellInfoForUncheck.ROW}.", LogLevel.WARN);
                            }
                        }
                        else
                        {
                            Log($"[{currentIp}] Invalid row index {cellInfoForUncheck.ROW} for IP {currentIp} when trying to uncheck checkbox.", LogLevel.WARN);
                        }
                    });
                }
                else
                {
                    Log($"[{currentIp}] Could not find CellData for IP {currentIp} to uncheck checkbox.", LogLevel.WARN);
                }
                #endregion
            }
            catch (Exception ex)
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    Log($"[{currentIp}] _ Error occurred: {ex.Message}", LogLevel.ERROR); //
                    SetFail(currentIp); //
                });
            }
            finally
            {
                onIpCompletedCallback?.Invoke();
            }
        }

        private async Task<bool> PatchInternalAsync(string ip, List<string> backupPathList, string remoteBackupPath, string remoteFolderPath, List<string> filesToCheck, string sourceDirectory, string processNameToCheck)
        {
            try
            {
                // 프로그램 실행 중 확인 및 종료 대기 (WaitForProcessToExitAsync로 변경)
                if (IsProgramRunning(ip, processNameToCheck)) //
                {
                    await Dispatcher.InvokeAsync(() => Log($"[{ip}] _ {processNameToCheck} is running : Waiting for exit ...")); //
                    await WaitForProcessToExitAsync(ip, processNameToCheck, 120); // 수정된 메서드 호출
                }
                else
                {
                    await Dispatcher.InvokeAsync(() => Log($"[{ip}] _ [{processNameToCheck}] is not running. Try to patch..")); //
                }

                // 백업 (CheckBackupFolderAsync로 변경)
                foreach (string strBackUpPath in backupPathList)
                {
                    await Dispatcher.InvokeAsync(() => Log($"[{ip}] _ [{strBackUpPath}] back up..")); //
                    if (!await CheckBackupFolderAsync(ip, strBackUpPath, remoteBackupPath)) // 수정된 메서드 호출
                    {
                        await Dispatcher.InvokeAsync(() => Log($"[{ip}] _ {strBackUpPath} backup failed or path doesn't exist", LogLevel.ERROR)); //
                        return false;
                    }
                    await Dispatcher.InvokeAsync(() => Log($"[{ip}] _ [{strBackUpPath}] Backup successfully!")); //
                }

                // 폴더 생성
                foreach (string folder in FoldersToCheck) // FoldersToCheck는 멤버 변수, 안정적이어야 함
                {
                    string fullFolderPath = System.IO.Path.Combine(remoteFolderPath, folder);
                    if (!Directory.Exists(fullFolderPath))
                    {
                        Directory.CreateDirectory(fullFolderPath); // 파일 시스템 작업, 예외 발생 가능성
                    }
                }

                // 파일 업데이트 (ReplaceFileAsync로 변경)
                int cnt = 0; //
                foreach (string fileName in filesToCheck) // filesToCheck는 파라미터로 전달
                {
                    cnt++; //
                    string localFilePath = System.IO.Path.Combine(sourceDirectory, fileName); //
                    string remoteFilePath = System.IO.Path.Combine(remoteFolderPath, fileName); //

                    bool updateNeeded = !System.IO.File.Exists(remoteFilePath) || IsFileUpdateNeeded(localFilePath, remoteFilePath); //

                    if (updateNeeded)
                    {
                        await Dispatcher.InvokeAsync(() => Log($"[{ip}] _ Updating: {fileName}."));
                        await ReplaceFileAsync(ip, remoteFolderPath, fileName, localFilePath); // 수정된 메서드 호출
                    }
                    else
                    {
                        await Dispatcher.InvokeAsync(() => Log($"[{ip}] _ Skip (up-to-date): {fileName}")); //
                    }
                    // 개별 파일 진행률 업데이트는 전체 진행률 계산 방식 변경으로 인해 여기서는 제거하거나 다르게 처리
                    // double currentIpFileProgress = (((double)cnt / (double)filesToCheck.Count)) * 100;
                    // await Dispatcher.InvokeAsync(() => {
                    //     // pbstatusBar.Value = currentIpFileProgress; //
                    //     // txtStatusBar.Text = currentIpFileProgress.ToString("0.0") + "%"; //
                    // });
                }
                return true;
            }
            catch (Exception ex)
            {
                await Dispatcher.InvokeAsync(() => Log($"[{ip}] Error in PatchInternalAsync: {ex.Message}", LogLevel.ERROR)); //
                return false;
            }
        }

        bool ProcessStart(string ip)
        {
            try
            {
                //func1
                ProcessStartInfo si = new ProcessStartInfo();
                si.FileName = "schtasks.exe";
                si.UseShellExecute = false;
                si.WindowStyle = ProcessWindowStyle.Hidden;
                si.CreateNoWindow = true;
                si.RedirectStandardInput = true;
                Process run = new Process();

                ResourceManager rscManager = new ResourceManager("AutoPatcher.Resource.UserInfo", typeof(MainWindow).Assembly);

                string id = rscManager.GetString($"{ModeType.ToUpper()}_{PCType.ToUpper()}_ID");
                string pw = rscManager.GetString($"{ModeType.ToUpper()}_{PCType.ToUpper()}_PW");

                si.Arguments = $"/run /tn {PCType.ToUpper()} /s {ip} /u {id} /p {pw}";                
                run.StartInfo = si;
                run.Start();
                Task.Delay(100);

                return true;

            }
            catch (Exception ex)
            {
                Log($"[{ip}] _ {ex.Message}", LogLevel.ERROR);
                return false;
            }
        }
        /*
            //    Thread.Sleep(300);

            //    // func2
            //    string remoteName = @"\\" + ip + @"\root\cimv2";

            //    ConnectionOptions con = new ConnectionOptions();
            //    ResourceManager rscManager = new ResourceManager("AutoPatcher.Resource.UserInfo", typeof(MainWindow).Assembly);

            //    con.Username = "bga";// rscManager.GetString($"{ModeType.ToUpper()}_{PCType.ToUpper()}_ID");
            //    con.Password = "vision";//rscManager.GetString($"{ModeType.ToUpper()}_{PCType.ToUpper()}_PW");

            //    ManagementScope managementScope = new ManagementScope(remoteName, con);
            //    managementScope.Options.Authentication = AuthenticationLevel.PacketPrivacy;
            //    managementScope.Connect();

            //    ManagementClass processClass = new ManagementClass(managementScope, new ManagementPath("Win32_Process"), null);

            //    ManagementBaseObject inParams = processClass.GetMethodParameters("Create");
            //    inParams["CommandLine"] = path;  

            //    ManagementBaseObject outParams = processClass.InvokeMethod("Create", inParams, null);

            //    uint returnCode = (uint)outParams["ReturnValue"];
            //    if (returnCode == 0)
            //    {
            //        Console.WriteLine("프로그램이 성공적으로 실행되었습니다.");
            //        return true;
            //    }
            //    else
            //    {
            //        Console.WriteLine("프로그램 실행 실패. 반환 코드: " + returnCode);
            //        return false;
            //    }
            //}
            //catch (UnauthorizedAccessException ex)
            //{
            //    Console.WriteLine($"액세스 오류: {ex.Message}");
            //    return false;
            //}
            //catch (System.Net.Sockets.SocketException ex)
            //{
            //    Console.WriteLine($"연결 오류: {ex.Message}");
            //    return false;
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"일반 오류: {ex.Message}");
            //    return false;
        */
        

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

        async Task WaitForProcessToExitAsync(string ip, string processName, int timeoutSeconds) //
        {
            int waitedSeconds = 0; //
                                   // IsProgramRunning은 WMI를 사용하므로 블로킹 호출일 수 있음. Task.Run으로 감싸는 것을 고려할 수 있으나,
                                   // 여기서는 빈번한 호출이 아니므로 일단 그대로 둠.
            while (IsProgramRunning(ip, processName)) //
            {
                if (waitedSeconds >= timeoutSeconds) //
                {
                    await Dispatcher.InvokeAsync(() => Log($"[{ip}] TimeOut: {processName} timeout.", LogLevel.ERROR)); //
                    break;
                }
                await Dispatcher.InvokeAsync(() => Log($"[{ip}] {processName} is running... waiting {waitedSeconds + 1}s..")); //
                await Task.Delay(100); // 비동기 지연
                waitedSeconds++; //
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
            Log($"[{ip}] _ Updated : {fileName}");
            Debug.WriteLine($"[{remoteFolderPath}] 업데이트 완료: {fileName}");
        }

        public async Task ReplaceFileAsync(string ip, string remoteFolderPath, string fileName, string localFilePath) //
        {
            string remoteFilePath = System.IO.Path.Combine(remoteFolderPath, fileName); //
            try
            {
                // 동기 파일 I/O 작업을 Task.Run으로 실행
                await Task.Run(() =>
                {
                    if (System.IO.File.Exists(remoteFilePath)) // 원격 파일 존재하면 삭제
                    {
                        System.IO.File.Delete(remoteFilePath); //
                    }
                    System.IO.File.Copy(localFilePath, remoteFilePath); //
                });
                await Dispatcher.InvokeAsync(() => Log($"[{ip}] _ Updated : {fileName}")); //
            }
            catch (Exception ex)
            {
                await Dispatcher.InvokeAsync(() => Log($"[{ip}] _ Failed to replace {fileName}: {ex.Message}", LogLevel.ERROR));
                throw; // 예외를 다시 던져서 PatchInternalAsync에서 처리하도록 함
            }
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
                    if (file.Contains("CAM") || file.Contains(".git"))
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

        public async Task<bool> CheckBackupFolderAsync(string ip, string src, string des) //
        {
            // ... (파라미터 유효성 검사 및 로그) ...
            try
            {
                if (!Directory.Exists(src)) //
                {
                    await Dispatcher.InvokeAsync(() => Log($"[{ip}] Source backup dir {src} doesn't exist", LogLevel.ERROR)); //
                    return false;
                }
                // BackupDirectory는 동기 I/O 작업을 포함하므로 Task.Run으로 실행
                await Task.Run(() => BackupDirectory(src, des)); //
                return true;
            }
            catch (Exception ex)
            {
                await Dispatcher.InvokeAsync(() => Log($"[{ip}] Error during backup of {src}: {ex.Message}", LogLevel.ERROR)); //
                return false;
            }
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
                if (!src.ToLower().Contains(".git") || !src.ToLower().Contains("git"))
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

        private async void btnSetPatchDirectory(object sender, RoutedEventArgs e)
        {
            // Use FolderBrowserDialog to select a folder
            if (string.IsNullOrEmpty(ProcessNameToCheck))
            {
                await Task.Delay(1);
                Log("Set the process name",LogLevel.ERROR);                
                System.Windows.MessageBox.Show("Set the process name");
                return;
            }

            CommonOpenFileDialog cofd = new CommonOpenFileDialog();

            cofd.IsFolderPicker = true;

            if(cofd.ShowDialog()==CommonFileDialogResult.Ok)
            {
                SourceDirectory = cofd.FileName + "\\";
                await Task.Delay(1);
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
                    System.Action action = delegate
                    {
                        Log($"Loading...{dlg.FileName}");
                        ExcelPath = dlg.FileName;
                        LoadExcelData(dlg.FileName);
                        lblCurExcel.Content = dlg.FileName;
                        Log($"Loaded machine list : {ModeType}");
                    };
                    this.Dispatcher.Invoke(action);
                }
            }
        }

        private async void Testfunc(object sender , RoutedEventArgs e)
        {
            //ProcessStart("55.60.231.166", "D:\\main\\bin\\HDSInspector.exe");
            Log("Start patch");

            if (SelectedMode == -1)
            {
                Log("Select a patch node", LogLevel.ERROR);
                System.Windows.MessageBox.Show("Select a patch node");
            }

            if (FilesToCheck.Count == 0)
            {
                Log("File count is 0");
            }

            for (int i = 0; i <= 100; i++)
            {
                pbstatusBar.Value = i;
                txtStatusBar.Text = i.ToString() + "%";
                Log($"Test..{i}");
                await Task.Delay(1);
            }


        }

        private async void StartPatch(object sender, RoutedEventArgs e)
        {
            Log("Start patch..");
            if (SourceDirectory == null)
            {
                for (int i = 0; i < 3; i++)
                {
                    await Task.Delay(500);
                    Log("No selected source directory", LogLevel.WARN);
                }
                System.Windows.MessageBox.Show("No selected source directory");
            }

            else
            {
                foreach (string file in FilesToCheck)
                {
                    Log($"name : {file}");
                    await Task.Delay(1);
                }

                SetFail("192.168.30.100");
                Log($"192.168.30.100 is fail");
                await Task.Delay(1000);
                SetComplete("192.168.30.110");
                Log($"192.168.30.110 is complete");
                await Task.Delay(1000);
                SetComplete("192.168.30.200");
                Log($"192.168.30.200 is complete");
            }

            await Task.Delay(1000);

            //await Testfunc();
        }

        private async void btnRunPatch_Click(object sender, RoutedEventArgs e)
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
                Log("No equipment specified", LogLevel.ERROR);
                System.Windows.MessageBox.Show("No selected mahcine");
                
                return;
            }

            Log($"Start Patching .. {IPAddresses.Count} machines selected");
            await Task.Delay(10);

            await StartAutoPatch();

            Log("All Patch has been done");
            //if (StartAutoPatch())
            //{
            //    Log("All Patch completed");
            //}
            //else
            //{
            //    Log("Failed to patch", LogLevel.WARN);
            //}
        }

        private T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);

            // 부모가 null이 아니고, 원하는 타입이 될 때까지 반복
            while (parent != null && !(parent is T))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }

            return parent as T;
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as System.Windows.Controls.CheckBox;
            var dataItem = checkBox.DataContext as RowData;
            
            string ip = "";

            var sp = (StackPanel)checkBox.Parent;
            var cell = FindParent<System.Windows.Controls.DataGridCell>(sp);

            var row = DataGrid.Items.IndexOf(cell.DataContext);
            var col = DataGrid.Columns.IndexOf(cell.Column);

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
                if (flag)
                {
                    IPAddresses.Add(ip);
                    CellDatas.Add(new CellData
                    {
                        IP = ip,
                        ROW = row,
                        COLUMN = col
                    });
                }
                else
                {
                    IPAddresses.Remove(ip);
                    CellData removeItem = CellDatas.FirstOrDefault(item => item.IP==ip);
                    CellDatas.Remove(removeItem);
                }
            }
            System.Action action = delegate
            {
                lblList.Content = $"Number of selected machine : {IPAddresses.Count}";
            };            
            this.Dispatcher.Invoke(action);
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
                    if(ModeType=="SII")
                    {
                        PCType = "Inline";
                    }
                    else
                    {
                        PCType = "Main";
                    }
                    _TypeArray[0] = true;                    
                    ProcessNameToCheck = "HDSInspector.exe";
                    lblProcName.Content = ProcessNameToCheck;
                }

                else if (btn.Name.Contains("Vision"))
                {
                    if(ModeType=="SII")
                    {
                        PCType = "Inline";
                    }
                    else
                    {
                        PCType = "Vision";
                    }
                    _TypeArray[1] = false;                    
                    ProcessNameToCheck = "IS.exe";
                    lblProcName.Content = ProcessNameToCheck;
                }

                #region 기타파일
                else if (btn.Name.Contains("Etc"))
                {
                    PCType = "etc";
                    ProcessNameToCheck = "etc";
                    lblProcName.Content = ProcessNameToCheck;
                    cmbFileType.Visibility = Visibility.Visible;
                }
                #endregion
                Log($"{PCType}_{ProcessNameToCheck} selected ");
            }
            #endregion

            
            
        }

        private void cmbType_Selected(object sender,RoutedEventArgs e)
        {
            ComboBoxItem item = sender as ComboBoxItem;
            bool isFile = (item.TabIndex == 0) ? true : false;
            if(isFile)
            {
                
            }
            else
            {

            }
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

    public class CellData
    {
        public string IP { get; set; }
        public int ROW { get; set; }
        public int COLUMN { get; set; }
    }
}

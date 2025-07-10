using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Runtime.CompilerServices;

namespace Updater
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Updater will start after 5s");
            for (int i=0; i<5; i++)
            {
                Console.WriteLine($"Waiting... {i+1}s / 5s");
                Thread.Sleep(1000);
            }            

            try
            {
                // 2. 경로 설정
                // AppContext.BaseDirectory는 현재 실행 중인 updater.exe의 위치 (예: D:\main\temp_update\)
                string stagingDir = AppContext.BaseDirectory;
                // Path.GetFullPath와 '..'을 사용하여 상위 폴더를 거쳐 실제 프로그램 경로를 찾습니다.
                string targetDir = Path.GetFullPath(Path.Combine(stagingDir, @"..\", "bin"));
                string configDir = Path.GetFullPath(Path.Combine(stagingDir, @"..\", "config"));
                string backupDir = Path.GetFullPath(Path.Combine(stagingDir, @"..\", "backup"));
                string processName = "";

                var dir = new DirectoryInfo(targetDir);
                foreach (FileInfo file in dir.GetFiles())
                {
                    if(file.Name.Contains("IS.exe"))
                    {
                        processName = "IS.exe";
                        break;
                    }
                    else if(file.Name.Contains("HDSInspector.exe"))
                    {
                        processName = "HDSInspector.exe";
                        break;
                    }
                }
                Console.WriteLine($"Process Name : {processName}");

                // 파일 복사 전에 대상 프로세스 종료
                if (!string.IsNullOrEmpty(processName))
                {
                    Log($"Attempting to kill process: {processName}");
                    KillProcess(processName);
                }

                // 백업리스트 설정
                List<string> backupList = new List<string>();
                backupList.Add(targetDir);
                if(processName == "HDSInspector.exe")
                {
                    backupList.Add(configDir);
                }
                
                string mainAppPath = Path.Combine(targetDir, processName);

                Console.WriteLine($"Staging Directory: {stagingDir}");
                Console.WriteLine($"Target Directory: {targetDir}");

                if (!Directory.Exists(targetDir))
                {
                    Log($"Target directory not found: {targetDir}");
                    return;
                }

                // 백업
                foreach(string backupItem in backupList)
                {
                    if (Directory.Exists(backupItem))
                    {
                        BackupDirectory(backupItem, backupDir);
                    }
                    else
                    {
                        Log($"Backup item not found: {backupItem}");
                    }
                }
                

                // update 폴더의 모든 파일과 하위 폴더를 실제 프로그램 폴더로 복사 (덮어쓰기)
                CopyDirectory(stagingDir, targetDir, true);

                Console.WriteLine("Update complete. Restarting application...");

                // 성공 플래그 파일 생성
                string successFlagFilePath = Path.Combine(targetDir, "update_success.flag");
                try
                {
                    File.Create(successFlagFilePath).Dispose(); // 파일 생성 및 즉시 닫기
                    Log($"Success flag file created: {successFlagFilePath}");
                }
                catch (Exception flagEx)
                {
                    Log($"Failed to create success flag file: {flagEx.Message}");
                }

                // 메인일 경우 실행
                if (processName== "HDSInspector.exe" && Directory.Exists(targetDir))
                {
                    string updaterPath = Path.Combine(targetDir, "HDSInspector.exe");

                    if (File.Exists(updaterPath))
                    {                        
                        Process.Start(updaterPath);
                    }
                }

                string command = $"/c timeout /t 2 /nobreak > nul && rd /s /q \"{stagingDir}\"";

                ProcessStartInfo psDeleteTemp = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Verb = "runas",
                    Arguments = command,
                    WindowStyle = ProcessWindowStyle.Hidden,    // 창을 숨김
                    CreateNoWindow = true,                      // 새 창을 만들지 않음
                    UseShellExecute = false
                };
                
                Process.Start(psDeleteTemp);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                // 예외 발생 시 로그 파일 생성
                Log($"An error occurred during update: {ex.ToString()}");

                // 실패 플래그 파일 생성
                string stagingDir = AppContext.BaseDirectory;
                string failureFlagFilePath = Path.Combine(targetDir, "update_failure.flag");
                try
                {
                    File.Create(failureFlagFilePath).Dispose(); // 파일 생성 및 즉시 닫기
                    Log($"Failure flag file created: {failureFlagFilePath}");
                }
                catch (Exception flagEx)
                {
                    Log($"Failed to create failure flag file: {flagEx.Message}");
                }

                // 콘솔창 유지
                Console.ReadLine();
            }
            return;
        }

        private static void BackupDirectory(string src, string dest)
        {
            if (string.IsNullOrEmpty(src) || string.IsNullOrEmpty(dest)) return;

            var dir = new DirectoryInfo(src);
            
            if (dir.GetFiles().Length == 0) return;

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

        private static void CopyFolder(string src, string dest)
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
                        if(fileName.Contains("updater.exe"))
                        {
                            // updater.exe 자기 자신은 복사 대상에서 제외
                            continue;
                        }
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

        /// <summary>
        /// 폴더를 재귀적으로 복사하는 헬퍼 메서드
        /// </summary>
        /// <param name="sourceDir">소스 폴더</param>
        /// <param name="destinationDir">대상 폴더</param>
        /// <param name="recursive">하위 폴더 포함 여부</param>
        private static void CopyDirectory(string sourceDir, string destinationDir, bool recursive)
        {
            var dir = new DirectoryInfo(sourceDir);
            if (!dir.Exists)
                throw new DirectoryNotFoundException($"Source directory not found: {dir.FullName}");

            DirectoryInfo[] dirs = dir.GetDirectories();
            Directory.CreateDirectory(destinationDir);

            foreach (FileInfo file in dir.GetFiles())
            {
                // updater.exe 자기 자신은 복사 대상에서 제외
                if (file.Name.Equals("updater.exe", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string targetFilePath = Path.Combine(destinationDir, file.Name);
                file.CopyTo(targetFilePath, true); // true 옵션으로 덮어쓰기 허용
            }

            if (recursive)
            {
                foreach (DirectoryInfo subDir in dirs)
                {
                    string newDestinationDir = Path.Combine(destinationDir, subDir.Name);
                    CopyDirectory(subDir.FullName, newDestinationDir, true);
                }
            }
        }

        // 프로세스를 종료하는 헬퍼 메서드 추가
        private static void KillProcess(string processName)
        {
            try
            {
                foreach (Process proc in Process.GetProcessesByName(processName.Replace(".exe", "")))
                {
                    proc.Kill();
                    Log($"Process {processName} killed successfully.");
                    proc.WaitForExit(5000); // 5초 대기
                }
            }
            catch (Exception ex)
            {
                Log($"Error killing process {processName}: {ex.Message}");
            }
        }

        /// <summary>
        /// 간단한 로그 파일 작성 메서드
        /// </summary>
        private static void Log(string message)
        {
            try
            {
                string logFilePath = Path.Combine(AppContext.BaseDirectory, "updater_log.txt");
                File.AppendAllText(logFilePath, $"[{DateTime.Now}] {message}{Environment.NewLine}");
            }
            catch
            {
                // 로깅 실패 시 무시
            }
        }
    }
}


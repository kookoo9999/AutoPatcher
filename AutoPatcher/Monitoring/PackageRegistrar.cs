using System;
using System.IO;

namespace AutoPatcher
{
    /// <summary>
    /// 중앙 배포 등록: package_{version} 폴더를 완성한 뒤에 version.txt를 갱신한다.
    /// 순서가 바뀌면 폴링 중인 PatchAgent가 "버전은 바뀌었는데 패키지가 없는" 상태를 만나 스테이징에 실패한다.
    /// </summary>
    internal static class PackageRegistrar
    {
        public static string PackageDir(string pcTypeRoot, string version) =>
            Path.Combine(pcTypeRoot, "package_" + version);

        public static string ReadServerVersion(string pcTypeRoot)
        {
            string versionFile = Path.Combine(pcTypeRoot, "version.txt");
            return File.Exists(versionFile) ? File.ReadAllText(versionFile).Trim() : string.Empty;
        }

        public static void Register(string pcTypeRoot, string version, string sourceDir, Action<string> log)
        {
            if (!Directory.Exists(sourceDir))
                throw new DirectoryNotFoundException($"패치 폴더를 찾을 수 없습니다: {sourceDir}");

            Directory.CreateDirectory(pcTypeRoot);

            string finalDir = PackageDir(pcTypeRoot, version);
            if (Directory.Exists(finalDir))
                throw new InvalidOperationException($"이미 존재하는 패키지입니다: {finalDir}");

            // 업로드 중인 폴더는 이름이 달라 에이전트가 절대 읽지 않는다. 완성 후 이름 변경으로 공개한다.
            string uploadingDir = finalDir + ".uploading";
            if (Directory.Exists(uploadingDir)) Directory.Delete(uploadingDir, recursive: true);

            log($"패키지 복사 시작: {sourceDir} → {uploadingDir}");
            long srcFiles = 0, srcBytes = 0;
            CopyDirectory(sourceDir, uploadingDir, ref srcFiles, ref srcBytes);

            long dstFiles = 0, dstBytes = 0;
            Measure(uploadingDir, ref dstFiles, ref dstBytes);
            if (srcFiles != dstFiles || srcBytes != dstBytes)
            {
                Directory.Delete(uploadingDir, recursive: true);
                throw new IOException($"복사 검증 실패 (원본 {srcFiles}개/{srcBytes}B, 대상 {dstFiles}개/{dstBytes}B)");
            }
            log($"복사 검증 완료: {dstFiles}개 파일, {dstBytes:N0} bytes");

            Directory.Move(uploadingDir, finalDir);
            WriteFileAtomic(Path.Combine(pcTypeRoot, "version.txt"), version);
            log($"version.txt 갱신 완료: {version}");
        }

        // PatchAgent는 패키지를 bin에 그대로 덮어쓰므로, Updater 전용 파일과 .git 폴더는 제외한다.
        private static bool IsExcludedFile(string name) =>
            name.Equals("updater.exe", StringComparison.OrdinalIgnoreCase);

        private static bool IsExcludedDir(string name) =>
            name.IndexOf(".git", StringComparison.OrdinalIgnoreCase) >= 0;

        private static void CopyDirectory(string src, string dst, ref long files, ref long bytes)
        {
            Directory.CreateDirectory(dst);
            foreach (var file in new DirectoryInfo(src).GetFiles())
            {
                if (IsExcludedFile(file.Name)) continue;
                file.CopyTo(Path.Combine(dst, file.Name), overwrite: true);
                files++;
                bytes += file.Length;
            }
            foreach (var dir in new DirectoryInfo(src).GetDirectories())
            {
                if (IsExcludedDir(dir.Name)) continue;
                CopyDirectory(dir.FullName, Path.Combine(dst, dir.Name), ref files, ref bytes);
            }
        }

        private static void Measure(string dir, ref long files, ref long bytes)
        {
            foreach (var file in new DirectoryInfo(dir).GetFiles())
            {
                files++;
                bytes += file.Length;
            }
            foreach (var sub in new DirectoryInfo(dir).GetDirectories())
                Measure(sub.FullName, ref files, ref bytes);
        }

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
                File.Copy(tmp, path, overwrite: true);
                File.Delete(tmp);
            }
        }
    }
}

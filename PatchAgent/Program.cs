using System;
using System.IO;
using System.Threading;

namespace PatchAgent
{
    internal class Program
    {
        static int Main(string[] args)
        {
            string baseDir = AppContext.BaseDirectory;
            var log = new Logger(baseDir);

            using (var mutex = new Mutex(initiallyOwned: false, name: "Global\\HDS_PatchAgent_SingleInstance", createdNew: out bool isNew))
            {
                bool acquired = false;
                try
                {
                    // 이전 사이클이 아직 실행 중이면(대용량 복사 등) 이번 실행은 건너뛴다.
                    acquired = mutex.WaitOne(TimeSpan.Zero, exitContext: false);
                    if (!acquired)
                    {
                        log.Warn("이전 PatchAgent 실행이 아직 진행 중입니다. 이번 사이클은 건너뜁니다.");
                        return 0;
                    }

                    string configPath = Path.Combine(baseDir, "PatchAgent.ini");
                    AgentConfig config = AgentConfig.Load(configPath);

                    var worker = new PatchWorker(config, log);
                    worker.RunCycle();
                    return 0;
                }
                catch (Exception ex)
                {
                    log.Error($"처리되지 않은 오류: {ex}");
                    return 1;
                }
                finally
                {
                    if (acquired) mutex.ReleaseMutex();
                }
            }
        }
    }
}

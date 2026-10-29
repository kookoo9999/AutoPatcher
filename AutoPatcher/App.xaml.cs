using System;
using System.Threading;
using System.Windows;

namespace AutoPatcher
{
    /// <summary>
    /// App.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class App : Application
    {
        private const string MutexName = "HDS_AutoPatcher_SingleInstance";
        private Mutex _instanceMutex;
        private bool _ownsMutex;

        protected override void OnStartup(StartupEventArgs e)
        {
            _instanceMutex = new Mutex(initiallyOwned: true, name: MutexName, createdNew: out bool createdNew);
            _ownsMutex = createdNew;

            if (!createdNew)
            {
                MessageBox.Show(
                    "프로그램이 이미 실행 중입니다.\n작업 표시줄 우측 하단의 시스템 트레이 아이콘을 확인해 주세요.",
                    "HDS Auto Patcher",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                Shutdown();
                return;
            }

            base.OnStartup(e);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            if (_instanceMutex != null)
            {
                if (_ownsMutex)
                {
                    try { _instanceMutex.ReleaseMutex(); }
                    catch (ApplicationException) { }
                }
                _instanceMutex.Dispose();
                _instanceMutex = null;
            }

            base.OnExit(e);
        }
    }
}

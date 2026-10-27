using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Resources;

namespace AutoPatcher
{
    /// <summary>트레이 아이콘(NotifyIcon): 창 닫기/최소화 시 트레이로 숨기고 더블클릭·우클릭 메뉴로 복원/종료한다.</summary>
    public partial class MainWindow
    {
        private NotifyIcon _trayIcon;
        private bool _exitRequested;
        private WindowState _lastWindowState = WindowState.Normal;

        private void InitTrayIcon()
        {
            var menu = new ContextMenuStrip();
            var openItem = new ToolStripMenuItem("열기") { Font = new Font(menu.Font, System.Drawing.FontStyle.Bold) };
            openItem.Click += (s, e) => RestoreFromTray();

            var exitItem = new ToolStripMenuItem("종료");
            exitItem.Click += (s, e) => ExitApplication();

            menu.Items.Add(openItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(exitItem);

            _trayIcon = new NotifyIcon
            {
                Icon = LoadTrayIcon(),
                Text = "HDS Auto Patcher",
                ContextMenuStrip = menu,
                Visible = true,
            };
            _trayIcon.DoubleClick += (s, e) => RestoreFromTray();

            StateChanged += MainWindow_StateChanged;
            Closing += MainWindow_Closing;
            Closed += MainWindow_Closed;
        }

        private static Icon LoadTrayIcon()
        {
            try
            {
                StreamResourceInfo res = System.Windows.Application.GetResourceStream(
                    new Uri("pack://application:,,,/icons8-driver-booster-128.ico"));
                if (res?.Stream != null)
                {
                    using (res.Stream)
                        return new Icon(res.Stream);
                }
            }
            catch
            {
                // 리소스 로드 실패 시 실행 파일 아이콘으로 대체
            }

            try
            {
                string exePath = Process.GetCurrentProcess().MainModule?.FileName;
                if (!string.IsNullOrEmpty(exePath))
                {
                    Icon associated = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
                    if (associated != null) return associated;
                }
            }
            catch
            {
                // 무시하고 기본 아이콘 반환
            }

            return SystemIcons.Application;
        }

        private void MainWindow_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
            {
                HideToTray();
            }
            else
            {
                _lastWindowState = WindowState;
            }
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (_exitRequested) return;

            e.Cancel = true;
            HideToTray();
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            DisposeTrayIcon();
        }

        private void HideToTray()
        {
            if (WindowState != WindowState.Minimized)
                _lastWindowState = WindowState;

            ShowInTaskbar = false;
            Hide();
        }

        private void RestoreFromTray()
        {
            Show();
            ShowInTaskbar = true;
            WindowState = _lastWindowState == WindowState.Minimized ? WindowState.Normal : _lastWindowState;
            Activate();
        }

        private void ExitApplication()
        {
            _exitRequested = true;
            DisposeTrayIcon();
            Close();
        }

        private void DisposeTrayIcon()
        {
            if (_trayIcon == null) return;
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _trayIcon = null;
        }
    }
}

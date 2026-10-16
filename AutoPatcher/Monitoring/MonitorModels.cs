using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace AutoPatcher
{
    public enum MonitorState
    {
        None,
        UpToDate,
        Mismatch,
        Pending,
        Failed,
        NoResponse,
        NotInstalled,
    }

    /// <summary>모니터링 그리드의 셀 하나(설비 PC 1대).</summary>
    public class MonitorCell : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _ip;
        private MonitorState _state;
        private string _text;
        private string _detail;

        public string Ip
        {
            get => _ip;
            set
            {
                _ip = value?.Trim();
                OnPropertyChanged(nameof(Ip));
                OnPropertyChanged(nameof(HasIp));
            }
        }

        public bool HasIp => !string.IsNullOrEmpty(_ip);

        public MonitorState State
        {
            get => _state;
            set
            {
                _state = value;
                OnPropertyChanged(nameof(State));
                OnPropertyChanged(nameof(StateLabel));
            }
        }

        /// <summary>셀 두 번째 줄. 일치/미확인은 표시하지 않는다.</summary>
        public string StateLabel
        {
            get
            {
                switch (_state)
                {
                    case MonitorState.Mismatch: return "불일치";
                    case MonitorState.Pending: return "적용대기";
                    case MonitorState.Failed: return "실패";
                    case MonitorState.NoResponse: return "무응답";
                    case MonitorState.NotInstalled: return "미설치";
                    default: return string.Empty;
                }
            }
        }

        public string Text { get => _text; set { _text = value; OnPropertyChanged(nameof(Text)); } }
        public string Detail { get => _detail; set { _detail = value; OnPropertyChanged(nameof(Detail)); } }

        private void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>PatchAgent가 중앙 status 폴더에 남긴 상태파일 1개.</summary>
    internal class AgentStatus
    {
        public string MachineName;
        public string PCType;
        public string[] IPs = new string[0];
        public string LocalVersion;
        public string ServerVersion;
        public string State;
        public string Error;
        public DateTime LastWriteUtc;
    }

    /// <summary>PCType 하나(예: Main, Vision, Inline\IS)의 중앙 상태 스냅샷.</summary>
    internal class MonitorSnapshot
    {
        public string Root;
        public string ServerVersion;
        public string Error;
        public Dictionary<string, AgentStatus> ByIp = new Dictionary<string, AgentStatus>(StringComparer.OrdinalIgnoreCase);
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace Common
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct NETRESOURCE
    {
        public uint dwScope;
        public uint dwType;
        public uint dwDisplyType;
        public uint dwUsage;
        public string IpLocalName;
        public string IpRemoteName;
        public string IpComment;
        public string IpProvider;
    }
    public static class NetworkDriveConnector
    {
        public static NETRESOURCE NetResorce = new NETRESOURCE();

        [DllImport("mpr.dll", CharSet = CharSet.Auto)]
        public static extern int WNetUseConnection(
            IntPtr hwndOwner,
            [MarshalAs(UnmanagedType.Struct)] ref NETRESOURCE IpNetResorece,
            string IpPassword,
            string IpUserID,
            uint dwFlags,
            StringBuilder IpAccessName,
            ref int IpBufferSize,
            out uint IpResult);

        [DllImport("mpr.dll", EntryPoint = "WNetUseConnection2", CharSet = CharSet.Auto)]
        public static extern int WNetUseConnection2A(string IpName, int dwFlags, int fForce);

        public static int TryConnectNetwork(string remotePath, string userID, string pwd)
        {
            int capacity = 1028;
            uint resultFlags = 0;
            uint flags = 0;
            StringBuilder sb = new StringBuilder(capacity);
            NetResorce.dwType = 1; //공유 디스크
            NetResorce.IpLocalName =  null;
            NetResorce.IpRemoteName = remotePath;
            NetResorce.IpProvider = null;

            int result = WNetUseConnection(IntPtr.Zero, ref NetResorce, pwd, @userID, flags, sb, ref capacity, out resultFlags);

            return result;
        }

        public static void DisconnectionNetwork()
        {
            WNetUseConnection2A(NetResorce.IpRemoteName, 1, 0);
        }

        public static bool TryConnectResult(int state)
        {
            bool result = true;

            if (state == 0)
            {
                result = true;
            }
            else
            {
                result = false;

                switch (state)
                {
                    case ERROR_CODE.NO_ERROR: break;
                    case ERROR_CODE.ERROR_NO_NET_OR_BAD_SERVER: break;
                    case ERROR_CODE.ERROR_ACCESS_DENIED: break;
                    case ERROR_CODE.ERROR_ALREADY_ASSIGNED: break;
                    case ERROR_CODE.ERROR_BAD_DEV_TYPE: break;
                    case ERROR_CODE.ERROR_BAD_DEVICE: break;
                    case ERROR_CODE.ERROR_BAD_NET_NAME: break;
                    case ERROR_CODE.ERROR_BAD_PROFILE: break;
                    case ERROR_CODE.ERROR_BAD_PROVIDER: break;
                    case ERROR_CODE.ERROR_BAD_BUSY: break;
                    case ERROR_CODE.ERROR_CANCELLED: break;
                    case ERROR_CODE.ERROR_CANNOT_OPEN_PROFILE: break;
                    case ERROR_CODE.ERROR_DEVICE_ALREADY_REMEMBERED: break;
                    case ERROR_CODE.ERROR_EXTENDED_ERROR: break;
                    case ERROR_CODE.ERROR_INVAILD_PASSWORD: break;
                    case ERROR_CODE.ERROR_NO_NET_OR_BAD_PATH: break;
                    case ERROR_CODE.ERROR_INVAILD_ADRESS: break;
                    case ERROR_CODE.ERROR_NETWORK_BUSY: break;
                    case ERROR_CODE.ERROR_UNEXP_NET_ERR: break;
                    case ERROR_CODE.ERROR_INVALID_PARAMETER: break;
                    case ERROR_CODE.ERROR_BAD_USER_OR_PASSWORD: break;
                    case ERROR_CODE.ERROR_MULTIPLE_CONNECTION: break;
                }
            }
            return result;
        }

        public class ERROR_CODE
        {
            public const int NO_ERROR = 0;
            public const int ERROR_NO_NET_OR_BAD_SERVER = 53;
            public const int ERROR_BAD_USER_OR_PASSWORD = 1326;
            public const int ERROR_ACCESS_DENIED = 5;
            public const int ERROR_ALREADY_ASSIGNED = 85;
            public const int ERROR_BAD_DEV_TYPE = 66;
            public const int ERROR_BAD_DEVICE = 1200;
            public const int ERROR_BAD_NET_NAME = 67;
            public const int ERROR_BAD_PROFILE = 1206;
            public const int ERROR_BAD_PROVIDER = 1204;
            public const int ERROR_BAD_BUSY = 170;
            public const int ERROR_CANCELLED = 1223;
            public const int ERROR_CANNOT_OPEN_PROFILE = 1205;
            public const int ERROR_DEVICE_ALREADY_REMEMBERED = 1202;
            public const int ERROR_EXTENDED_ERROR = 1208;
            public const int ERROR_INVAILD_PASSWORD = 86;
            public const int ERROR_NO_NET_OR_BAD_PATH = 1203;
            public const int ERROR_INVAILD_ADRESS = 487;
            public const int ERROR_NETWORK_BUSY = 54;
            public const int ERROR_UNEXP_NET_ERR = 59;
            public const int ERROR_INVALID_PARAMETER = 87;
            public const int ERROR_MULTIPLE_CONNECTION = 1219;
        }
    }
}

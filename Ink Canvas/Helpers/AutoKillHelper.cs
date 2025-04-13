using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace Ink_Canvas.Helpers
{
    class AutoKillHelper
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint WM_CLOSE = 0x0010;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public static void KillEasiNoteFloatBall()
        {
            EnumWindows(new EnumWindowsProc((hWnd, lParam) =>
            {
                StringBuilder windowText = new StringBuilder(256);
                GetWindowText(hWnd, windowText, 256);
                if (string.IsNullOrEmpty(windowText.ToString()))
                {
                    GetWindowRect(hWnd, out RECT rect);
                    int height = rect.Bottom - rect.Top;
                    if (height <= 200 && rect.Left <= -175)
                    {
                        GetWindowThreadProcessId(hWnd, out uint processId);
                        Process process = Process.GetProcessById((int)processId);
                        if (process.ProcessName.Equals("EasiNote", StringComparison.OrdinalIgnoreCase))
                        {
                            PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                        }
                    }
                }
                return true;
            }), IntPtr.Zero);
        }
    }
}

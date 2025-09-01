using Autodesk.Revit.UI;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowsFormsApp1.Class
{
    public class DoubleClickTracker
    {
        private const int WH_MOUSE_LL = 14;
        private static readonly LowLevelMouseProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        private static readonly Stopwatch _stopwatch = new Stopwatch();
        private static readonly System.Timers.Timer _resetDoubleClickTimer = new System.Timers.Timer(300);

        public static void Start()
        {
            try
            {
                if (_hookID == IntPtr.Zero)
                {
                    _hookID = SetHook(_proc);
                    _resetDoubleClickTimer.Interval = SystemInformation.DoubleClickTime;
                    _resetDoubleClickTimer.Elapsed += (s, e) =>
                    {
                        GlobalSettings.IsDoubleClick = false;
                        _resetDoubleClickTimer.Stop();
                    };
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.Message);
            }
        }

        public static void Stop()
        {
            if (_hookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
            }
        }

        private static IntPtr SetHook(LowLevelMouseProc proc)
        {
            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                IntPtr foregroundWindow = GetForegroundWindow();
                GetWindowThreadProcessId(foregroundWindow, out uint processId);

                uint currentProcessId = (uint)Process.GetCurrentProcess().Id;
                if (processId == currentProcessId)
                {
                    const int WM_LBUTTONDOWN = 0x0201;

                    if (wParam == (IntPtr)WM_LBUTTONDOWN)
                    {
                        if (_stopwatch.IsRunning && _stopwatch.ElapsedMilliseconds < SystemInformation.DoubleClickTime)
                        {
                            _stopwatch.Reset();
                            GlobalSettings.IsDoubleClick = true;
                            _resetDoubleClickTimer.Stop();
                        }
                        else
                        {
                            _stopwatch.Restart();
                            _resetDoubleClickTimer.Start();
                        }
                    }
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
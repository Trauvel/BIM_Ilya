using Autodesk.Revit.UI;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowsFormsApp1.Class
{
    public class DoubleClickTracker : IDisposable
    {
        private const int WH_MOUSE_LL = 14;
        private static readonly LowLevelMouseProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        private static readonly Stopwatch _stopwatch = new Stopwatch();
        private static readonly System.Timers.Timer _resetDoubleClickTimer = new System.Timers.Timer(300);
        
        // Кэшируем ID процесса Revit для оптимизации
        private static uint _revitProcessId = 0;
        private static bool _isRevitActive = false;
        private static bool _isDisposed = false;

        // Добавляем финализатор для гарантированной очистки
        ~DoubleClickTracker()
        {
            Dispose(false);
        }

        public static void Start()
        {
            try
            {
                if (_hookID == IntPtr.Zero && !_isDisposed)
                {
                    // Кэшируем ID процесса Revit один раз
                    _revitProcessId = (uint)Process.GetCurrentProcess().Id;
                    
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
            try
            {
                if (_hookID != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_hookID);
                    _hookID = IntPtr.Zero;
                }
                
                if (_resetDoubleClickTimer != null)
                {
                    _resetDoubleClickTimer.Stop();
                    _resetDoubleClickTimer.Dispose();
                }
                
                _revitProcessId = 0;
                _isRevitActive = false;
                _isDisposed = true;
            }
            catch (Exception ex)
            {
                // Логируем ошибку, но не показываем диалог при закрытии
                System.Diagnostics.Debug.WriteLine($"Ошибка при остановке DoubleClickTracker: {ex.Message}");
            }
        }

        private static IntPtr SetHook(LowLevelMouseProc proc)
        {
            try
            {
                using (var curProcess = Process.GetCurrentProcess())
                using (var curModule = curProcess.MainModule)
                {
                    return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка при установке хука: {ex.Message}");
                return IntPtr.Zero;
            }
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            // Быстрый выход если код < 0 или хук не установлен
            if (nCode < 0 || _hookID == IntPtr.Zero || _isDisposed)
                return CallNextHookEx(_hookID, nCode, wParam, lParam);

            try
            {
                const int WM_LBUTTONDOWN = 0x0201;
                
                // Проверяем только события левой кнопки мыши
                if (wParam != (IntPtr)WM_LBUTTONDOWN)
                    return CallNextHookEx(_hookID, nCode, wParam, lParam);

                // Проверяем активность Revit только при нажатии левой кнопки
                if (!_isRevitActive)
                {
                    IntPtr foregroundWindow = GetForegroundWindow();
                    GetWindowThreadProcessId(foregroundWindow, out uint processId);
                    _isRevitActive = (processId == _revitProcessId);
                }

                // Если Revit не активен, пропускаем обработку
                if (!_isRevitActive)
                    return CallNextHookEx(_hookID, nCode, wParam, lParam);

                // Обрабатываем двойной клик
                if (_stopwatch.IsRunning && _stopwatch.ElapsedMilliseconds < SystemInformation.DoubleClickTime)
                {
                    _stopwatch.Reset();
                    GlobalSettings.IsDoubleClick = true;
                    _resetDoubleClickTimer?.Stop();
                }
                else
                {
                    _stopwatch.Restart();
                    _resetDoubleClickTimer?.Start();
                }
            }
            catch (Exception ex)
            {
                // Логируем ошибку, но не прерываем работу хука
                System.Diagnostics.Debug.WriteLine($"Ошибка в HookCallback: {ex.Message}");
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        // Метод для сброса флага активности Revit (вызывать при переключении окон)
        public static void ResetRevitActiveFlag()
        {
            _isRevitActive = false;
        }

        // Реализация IDisposable для гарантированной очистки
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                Stop();
            }
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
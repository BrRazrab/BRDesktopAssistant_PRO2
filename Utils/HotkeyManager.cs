using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Forms;

namespace BRDesktopAssistant.Utils
{
    [Flags]
    public enum HotkeyModifiers : uint
    {
        MOD_ALT = 0x0001,
        MOD_CONTROL = 0x0002,
        MOD_SHIFT = 0x0004,
        MOD_WIN = 0x0008
    }

    public static class HotkeyManager
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private static int _hotkeyId = 9001;
        private static Action? _onHotkey;

        public static void Register(Window window, HotkeyModifiers modifiers, Keys key, Action onHotkey)
        {
            _onHotkey = onHotkey;
            var helper = new WindowInteropHelper(window);
            var handle = helper.Handle;
            var source = HwndSource.FromHwnd(handle);
            source.AddHook(HwndHook);
            RegisterHotKey(handle, _hotkeyId, (uint)modifiers, (uint)key);
        }

        public static void Unregister(Window window)
        {
            var handle = new WindowInteropHelper(window).Handle;
            UnregisterHotKey(handle, _hotkeyId);
        }

        public static void Trigger() => _onHotkey?.Invoke();

        private static IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_HOTKEY = 0x0312;
            if (msg == WM_HOTKEY && wParam.ToInt32() == _hotkeyId)
            {
                _onHotkey?.Invoke();
                handled = true;
            }
            return IntPtr.Zero;
        }
    }

    public static class StartupHelper
    {
        private const string RunKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";

        public static void SetAutorun(bool enable)
        {
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunKey, true);
                var exe = System.Diagnostics.Process.GetCurrentProcess().MainModule!.FileName!;
                var name = "BRDesktopAssistant";
                if (enable) key?.SetValue(name, $""{exe}"");
                else key?.DeleteValue(name, false);
            }
            catch { }
        }
    }
}

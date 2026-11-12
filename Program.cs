using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Timers;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;

namespace LenovoBacklightAuto
{
    class Program
    {
        private static System.Timers.Timer _timer;
        private static IntPtr _keyboardHook;
        private static readonly KeyboardHookCallback _keyboardHookCallback = OnKeyboardHook;
        private const int WhKeyboardLl = 13;
        private const int WmKeyDown = 0x0100;
        private const int WmSysKeyDown = 0x0104;
        private const uint InjectedKeyFlags = 0x10 | 0x02;
        private static bool _isBacklightOn = false;
        private static bool _isTimerActive = true;
        private static object _keyboardControl;
        private static NotifyIcon _notifyIcon;
        static ToolStripMenuItem _activeItem;
        static ToolStripMenuItem[] _timeoutItems;
        static ToolStripMenuItem[] _intensityItems;
        static int _timeoutDuration = 600000; // Default 10 minutes
        static readonly int[] _timeoutOptions = { 5000, 60000, 120000, 300000, 600000, 900000, 1800000 };
        static readonly int[] _intensityOptions = { 1, 2 };
        private static bool _isBacklightManuallyOff = false;
        private static bool _isTempDisabled = false;
        private static readonly object _stateLock = new object();
        private static readonly Stopwatch _idleTime = Stopwatch.StartNew();
        private static MethodInfo _setBacklightStatusMethod;
        private static MethodInfo _getBacklightStatusMethod;
        private static bool _backlightErrorReported;

        [STAThread]
        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += ResolveEmbeddedAssembly;
            var coreAssembly = Assembly.Load("Keyboard_Core");
            var keyboardControlType = coreAssembly.GetType("Keyboard_Core.KeyboardControl", true);
            _keyboardControl = Activator.CreateInstance(keyboardControlType);
            _setBacklightStatusMethod = keyboardControlType.GetMethod("SetKeyboardBackLightStatus");
            _getBacklightStatusMethod = keyboardControlType.GetMethod("GetKeyboardBackLightStatus");
            if (_setBacklightStatusMethod == null || _getBacklightStatusMethod == null)
            {
                throw new MissingMethodException("The keyboard library must support backlight writes and status readback.");
            }

            _timer = new System.Timers.Timer(1000);
            _timer.Elapsed += OnTimerElapsed;
            _timer.AutoReset = true;

            _notifyIcon = new NotifyIcon();
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Lenovo_Keyboard_Backlight_Control.LenovoBacklightAuto2.ico"))
            {
                _notifyIcon.Icon = new Icon(stream);
            }
            _notifyIcon.Visible = true;
            _notifyIcon.Text = "Lenovo Backlight Auto";

            var contextMenu = new ContextMenuStrip();

            // Initialize intensity menu items
            var intensityMenu = new ToolStripMenuItem("Intensity");
            _intensityItems = new ToolStripMenuItem[_intensityOptions.Length];
            for (int i = 0; i < _intensityOptions.Length; i++)
            {
                int intensityValue = _intensityOptions[i];
                _intensityItems[i] = new ToolStripMenuItem($"{intensityValue}", null, (sender, e) => SetIntensity(intensityValue))
                {
                    Checked = intensityValue == 2
                };
                intensityMenu.DropDownItems.Add(_intensityItems[i]);
            }
            contextMenu.Items.Add(intensityMenu);

            // Initialize timeout menu items
            var timeoutMenu = new ToolStripMenuItem("Timeout Duration");
            _timeoutItems = new ToolStripMenuItem[_timeoutOptions.Length];
            for (int i = 0; i < _timeoutOptions.Length; i++)
            {
                int timeoutValue = _timeoutOptions[i];
                string timeoutText = GetTimeoutText(timeoutValue);
                _timeoutItems[i] = new ToolStripMenuItem(timeoutText, null, (sender, e) =>
                {
                    SetTimeout(timeoutValue);
                    foreach (var item in _timeoutItems)
                    {
                        item.Checked = item == (ToolStripMenuItem)sender;
                    }
                })
                {
                    Checked = timeoutValue == _timeoutDuration
                };
                timeoutMenu.DropDownItems.Add(_timeoutItems[i]);
            }
            contextMenu.Items.Add(timeoutMenu);

            // Add other menu items
            contextMenu.Items.Add(new ToolStripMenuItem("Turn On Backlight", null, (sender, e) => TurnOnBacklightManually()));
            contextMenu.Items.Add(new ToolStripMenuItem("Turn Off Backlight", null, (sender, e) => TurnOffBacklightManually()));
            _activeItem = new ToolStripMenuItem("Deactivate Timer", null, (sender, e) => ToggleTimer());
            contextMenu.Items.Add(_activeItem);
            contextMenu.Items.Add(new ToolStripMenuItem("Temporary Disable", null, (sender, e) => TempDisableBacklight()));
            contextMenu.Items.Add(new ToolStripMenuItem("Exit", null, (sender, e) => ExitApplication()));

            _notifyIcon.ContextMenuStrip = contextMenu;

            _keyboardHook = SetWindowsHookEx(WhKeyboardLl, _keyboardHookCallback, GetModuleHandle(null), 0);
            if (_keyboardHook == IntPtr.Zero)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not install the keyboard activity hook.");
            }
            SystemEvents.PowerModeChanged += SystemEvents_PowerModeChanged;

            _idleTime.Restart();
            _timer.Start();
            Application.Run();
        }

        static void SystemEvents_PowerModeChanged(object sender, PowerModeChangedEventArgs e)
        {
            if (e.Mode == PowerModes.Resume)
            {
                lock (_stateLock)
                {
                    _idleTime.Restart();
                    _isTempDisabled = false;
                    SetBacklightLevel(_isBacklightManuallyOff ? 0 : GetSelectedIntensity());
                }
            }
        }

        static string GetTimeoutText(int timeoutValue)
        {
            int seconds = timeoutValue / 1000;
            if (seconds < 60)
            {
                return $"{seconds}s";
            }
            else if (seconds < 3600)
            {
                int minutes = seconds / 60;
                return $"{minutes}m";
            }
            else
            {
                int hours = seconds / 3600;
                return $"{hours}h";
            }
        }

        static int GetSelectedIntensity()
        {
            return int.Parse(Array.Find(_intensityItems, item => item.Checked).Text);
        }

        static void SetIntensity(int intensity)
        {
            lock (_stateLock)
            {
                foreach (var item in _intensityItems)
                {
                    item.Checked = item.Text == $"{intensity}";
                }
                if (_isBacklightManuallyOff)
                {
                    TurnOnBacklightManually();
                }
                else
                {
                    _isTempDisabled = false;
                    _idleTime.Restart();
                    SetBacklightLevel(intensity);
                }
            }
        }

        static void TempDisableBacklight()
        {
            lock (_stateLock)
            {
                _isTempDisabled = true;
                SetBacklightLevel(0);
            }
        }

        static void SetTimeout(int duration)
        {
            lock (_stateLock)
            {
                _timeoutDuration = duration;
                _isTimerActive = true;
                _activeItem.Text = "Deactivate Timer";
                _idleTime.Restart();
            }
        }

        static void TurnOffBacklightManually()
        {
            lock (_stateLock)
            {
                _isBacklightManuallyOff = true;
                SetBacklightLevel(0);
            }
        }

        static void TurnOnBacklightManually()
        {
            lock (_stateLock)
            {
                _isBacklightManuallyOff = false;
                _isTempDisabled = false;
                _idleTime.Restart();
                SetBacklightLevel(GetSelectedIntensity());
            }
        }

        static Assembly ResolveEmbeddedAssembly(object sender, ResolveEventArgs args)
        {
            string resourceName = "Lenovo_Keyboard_Backlight_Control.Keyboard_Core.dll";
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                byte[] assemblyData = new byte[stream.Length];
                stream.Read(assemblyData, 0, assemblyData.Length);
                return Assembly.Load(assemblyData);
            }
        }

        private delegate IntPtr KeyboardHookCallback(int code, IntPtr message, IntPtr data);

        [StructLayout(LayoutKind.Sequential)]
        private struct KeyboardHookData
        {
            public uint VirtualKey;
            public uint ScanCode;
            public uint Flags;
            public uint Time;
            public UIntPtr ExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int hookId, KeyboardHookCallback callback, IntPtr module, uint threadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr message, IntPtr data);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string moduleName);

        static IntPtr OnKeyboardHook(int code, IntPtr message, IntPtr data)
        {
            if (code >= 0)
            {
                var key = Marshal.PtrToStructure<KeyboardHookData>(data);
                HandleKeyboardInput(message.ToInt32(), key.Flags);
            }

            return CallNextHookEx(_keyboardHook, code, message, data);
        }

        static void HandleKeyboardInput(int message, uint flags)
        {
            // Ignore software-generated input (including Caffeine), but let Windows receive it.
            if ((message == WmKeyDown || message == WmSysKeyDown) && (flags & InjectedKeyFlags) == 0)
            {
                HandleActivity();
            }
        }

        static void HandleActivity()
        {
            lock (_stateLock)
            {
                _idleTime.Restart();
                bool wasTempDisabled = _isTempDisabled;
                _isTempDisabled = false;
                if (!_isBacklightManuallyOff && (!_isBacklightOn || wasTempDisabled || _backlightErrorReported))
                {
                    SetBacklightLevel(GetSelectedIntensity());
                }
            }
        }

        static void OnTimerElapsed(object source, ElapsedEventArgs e)
        {
            lock (_stateLock)
            {
                if (_isBacklightManuallyOff || _isTempDisabled ||
                    (_isTimerActive && _idleTime.ElapsedMilliseconds >= _timeoutDuration))
                {
                    SetBacklightLevel(0);
                }
            }
        }

        static bool IsBacklightAtLevel(int level)
        {
            try
            {
                object[] arguments = { -1, null };
                uint result = (uint)_getBacklightStatusMethod.Invoke(_keyboardControl, arguments);
                return result == 0 && (int)arguments[0] == level;
            }
            catch (TargetInvocationException ex)
            {
                ReportBacklightError("Could not read the backlight state: " + (ex.InnerException ?? ex).Message);
                return false;
            }
        }

        static void SetBacklightLevel(int level)
        {
            try
            {
                // Redundant writes refresh Lenovo's OSD; read failures must not prevent a write.
                if (!IsBacklightAtLevel(level))
                {
                    _setBacklightStatusMethod.Invoke(_keyboardControl, new object[] { level, null });
                    if (!IsBacklightAtLevel(level))
                    {
                        ReportBacklightError("The keyboard did not confirm backlight level " + level + ".");
                        return;
                    }
                }

                _isBacklightOn = level != 0;
                _backlightErrorReported = false;
            }
            catch (TargetInvocationException ex)
            {
                ReportBacklightError((ex.InnerException ?? ex).Message);
            }
        }

        static void ReportBacklightError(string message)
        {
            Trace.TraceError("Keyboard backlight control failed: {0}", message);
            if (!_backlightErrorReported && _notifyIcon != null)
            {
                _notifyIcon.ShowBalloonTip(5000, "Lenovo Backlight Auto",
                    "Backlight control failed. Automatic switch-off will keep retrying. " + message,
                    ToolTipIcon.Warning);
            }
            _backlightErrorReported = true;
        }

        static void ToggleTimer()
        {
            lock (_stateLock)
            {
                _isTimerActive = !_isTimerActive;
                _activeItem.Text = _isTimerActive ? "Deactivate Timer" : "Activate Timer";
                if (_isTimerActive) _idleTime.Restart();
            }
        }

        static void ExitApplication()
        {
            lock (_stateLock)
            {
                _timer.Stop();
                _timer.Elapsed -= OnTimerElapsed;
                _timer.Dispose();
                SystemEvents.PowerModeChanged -= SystemEvents_PowerModeChanged;
                if (_keyboardHook != IntPtr.Zero && !UnhookWindowsHookEx(_keyboardHook))
                {
                    Trace.TraceError("Could not remove the keyboard activity hook: {0}",
                        new Win32Exception(Marshal.GetLastWin32Error()).Message);
                }
                _keyboardHook = IntPtr.Zero;
                _notifyIcon.Dispose();
                Environment.Exit(0);
            }
        }
    }
}

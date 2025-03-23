using System;
using System.Drawing;
using System.Reflection;
using System.Timers;
using System.Windows.Forms;
using System.IO;
using Gma.System.MouseKeyHook;
using Microsoft.Win32;

namespace LenovoBacklightAuto
{
    class Program
    {
        private static System.Timers.Timer _timer;
        private static IKeyboardMouseEvents _globalHook;
        private static bool _isBacklightOn = false;
        private static bool _isTimerActive = true;
        private static object _keyboardControl;
        private static NotifyIcon _notifyIcon;
        static ToolStripMenuItem _activeItem;
        static ToolStripMenuItem[] _timeoutItems;
        static ToolStripMenuItem[] _intensityItems;
        static int _timeoutDuration = 300000; // Default 5 minutes
        static readonly int[] _timeoutOptions = { 60000, 120000, 300000, 600000, 900000, 1800000 };
        static readonly int[] _intensityOptions = { 1, 2 };
        private static bool _isBacklightManuallyOff = false;
        private static bool _isTempDisabled = false;

        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += ResolveEmbeddedAssembly;
            var coreAssembly = Assembly.Load("Keyboard_Core");
            var keyboardControlType = coreAssembly.GetType("Keyboard_Core.KeyboardControl");
            _keyboardControl = Activator.CreateInstance(keyboardControlType);

            _timer = new System.Timers.Timer(_timeoutDuration);
            _timer.Elapsed += OnTimerElapsed;
            _timer.AutoReset = false;

            _globalHook = Hook.GlobalEvents();
            _globalHook.MouseDownExt += GlobalHookMouseDownExt;
            _globalHook.KeyDown += GlobalHookKeyDown;

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

            SystemEvents.PowerModeChanged += SystemEvents_PowerModeChanged;

            _timer.Start();
            Application.Run();
        }

        static void SystemEvents_PowerModeChanged(object sender, PowerModeChangedEventArgs e)
        {
            if (e.Mode == PowerModes.Resume)
            {
                if (!_isBacklightOn && !_isBacklightManuallyOff && !_isTempDisabled)
                {
                    SetBacklightLevel(2);
                    _isBacklightOn = true;
                }
                else if (_isTempDisabled)
                {
                    _isTempDisabled = false;
                    SetBacklightLevel(2);
                    _isBacklightOn = true;
                    if (_isTimerActive) _timer.Start();
                }
                if (_isTimerActive && !_isBacklightManuallyOff)
                {
                    _timer.Stop();
                    _timer.Start();
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

        static void SetIntensity(int intensity)
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
                SetBacklightLevel(intensity);
            }
        }

        static void TempDisableBacklight()
        {
            SetBacklightLevel(0);
            _isTempDisabled = true;
            _timer.Stop();
        }

        static void SetTimeout(int duration)
        {
            _timeoutDuration = duration;
            _timer.Interval = _timeoutDuration;
        }

        static void TurnOffBacklightManually()
        {
            SetBacklightLevel(0);
            _isBacklightOn = false;
            _isBacklightManuallyOff = true;
        }

        static void TurnOnBacklightManually()
        {
            int selectedIntensity = Array.Find(_intensityItems, item => item.Checked).Text[0] - '0';
            SetBacklightLevel(selectedIntensity);
            _isBacklightOn = true;
            _isBacklightManuallyOff = false;

            // Deactivate the timer and update the menu item text
            _isTimerActive = false;
            _activeItem.Text = "Activate Timer";
            _timer.Stop();
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

        static void GlobalHookMouseDownExt(object sender, MouseEventExtArgs e) => HandleActivity();
        static void GlobalHookKeyDown(object sender, KeyEventArgs e) => HandleActivity();

        static void HandleActivity()
        {
            if (!_isBacklightOn && !_isBacklightManuallyOff && !_isTempDisabled)
            {
                SetBacklightLevel(2);
                _isBacklightOn = true;
            }
            else if (_isTempDisabled)
            {
                _isTempDisabled = false;
                SetBacklightLevel(2);
                _isBacklightOn = true;
                if (_isTimerActive) _timer.Start();
            }
            if (_isTimerActive)
            {
                _timer.Stop();
                _timer.Start();
            }
        }

        static void OnTimerElapsed(object source, ElapsedEventArgs e)
        {
            SetBacklightLevel(0);
            _isBacklightOn = false;
        }

        static void SetBacklightLevel(int level)
        {
            try
            {
                var setBacklightStatusMethod = _keyboardControl.GetType().GetMethod("SetKeyboardBackLightStatus",
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic);
                setBacklightStatusMethod.Invoke(_keyboardControl, new object[] { level, null });
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        static void ToggleTimer()
        {
            _isTimerActive = !_isTimerActive;
            _activeItem.Text = _isTimerActive ? "Deactivate Timer" : "Activate Timer";
            if (_isTimerActive) _timer.Start();
            else _timer.Stop();
        }

        static void ExitApplication()
        {
            _globalHook.Dispose();
            _notifyIcon.Dispose();
            Environment.Exit(0);
        }
    }
}

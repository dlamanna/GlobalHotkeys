using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using CSCore.CoreAudioAPI;

namespace GlobalHotkeys
{
    public partial class Form1 : Form
    {
        private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        private const int WM_APPCOMMAND = 0x319;
        private const int WM_CHAR = 0x0102;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private Boolean isPetBattling = false;
        public const int VK_MEDIA_PLAY_PAUSE = 0xB3;
        public const int KEYEVENTF_KEYDOWN = 0x0001;        //Key down flag
        public const int KEYEVENTF_KEYUP = 0x0002;          //Key up flag
        KeyboardHook hook = new KeyboardHook();

        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);
        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, uint wParam, uint lParam);
        //[DllImport("user32.dll")]
        //internal static extern uint SendInput(uint nInputs, [MarshalAs(UnmanagedType.LPArray), In] INPUT[] pInputs,int cbSize);

        public Form1()
        {
            InitializeComponent();

            // register the event that is fired after the key press.
            hook.KeyPressed += new EventHandler<KeyPressedEventArgs>(hook_KeyPressed);

            /// Hotkey List
            // Mute/Unmute PUBG Toggle
            hook.RegisterHotKey((ModifierKeys)2, Keys.NumPad1);

            // Pause/Unpause Spotify Toggle
            hook.RegisterHotKey((ModifierKeys)2, Keys.NumPad2);

            // Mute/Unmute Teamspeak Toggle
            hook.RegisterHotKey((ModifierKeys)2, Keys.NumPad3);

            // Toggle Pet Battle (A-key spam)
            hook.RegisterHotKey((ModifierKeys)2, Keys.NumPad4);

            // Mute/Unmute BDO Toggle
            hook.RegisterHotKey((ModifierKeys)2, Keys.NumPad5);

            // Mute/Unmute WoW Toggle
            hook.RegisterHotKey((ModifierKeys)2, Keys.NumPad6);
        }

        void hook_KeyPressed(object sender, KeyPressedEventArgs e)
        {
            Console.WriteLine("Modifier: {0}({1}),\tKey: {2}({3})", e.Modifier.ToString(), (int)e.Modifier, e.Key.ToString(), (int)e.Key);

            String nirCMDPath = @"C:\Windows\SysWOW64\nircmd.exe";
            String nirCMDCommand = "muteappvolume";
            String exeName = "";
            String processName = "TslGame";


            Thread t = new Thread(setVolume);
            //Thread s = new Thread(petBattle);
            t.SetApartmentState(ApartmentState.MTA);
            //s.SetApartmentState(ApartmentState.MTA);
            Console.WriteLine("!!! After setting apartment state: {0}", t.GetApartmentState());

            switch ((int)e.Key)
            {
                case 97:
                    Console.WriteLine("!!! Do PUBG Toggle");
                    //exeName = " TslGame.exe ";
                    Process[] PUBGProcess = Process.GetProcessesByName(processName);
                    if (PUBGProcess.Length == 1)
                        t.Start();
                    else if (PUBGProcess.Length > 1)
                    {
                        t.Start();
                        Console.WriteLine("### More than 1 PUBG process found");
                    }
                    else
                        Console.WriteLine("### No PUBG process found");
                    break;
                case 98:
                    Console.WriteLine("!!! Do Spotify Toggle");
                    keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_KEYDOWN, 0);
                    keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_KEYUP, 0);
                    break;
                case 99:
                    Console.WriteLine("!!! Do Teamspeak Toggle");
                    exeName = " ts3client_win64.exe ";
                    nirCMDCommand += (exeName + "2");                               // the 2 is a toggle flag
                    runProgram(nirCMDPath, nirCMDCommand);
                    break;
                /*case 100:
                    Console.WriteLine("!!! Do PetBattle Toggle");
                    isPetBattling = !isPetBattling;
                    if (isPetBattling) s.Start();
                    else s.Abort();
                    break;*/
                case 101:
                    Console.WriteLine("!!! Do BlackDesert Toggle");
                    exeName = " BlackDesert64.exe ";
                    nirCMDCommand += (exeName + "2");                               // the 2 is a toggle flag
                    runProgram(nirCMDPath, nirCMDCommand);
                    break;
                case 102:
                    Console.WriteLine("!!! Do WoW Toggle");
                    exeName = " Wow.exe ";
                    nirCMDCommand += (exeName + "2");                               // the 2 is a toggle flag
                    runProgram(nirCMDPath, nirCMDCommand);
                    break;
                default:
                    break;
            }

        }

        private AudioSessionManager2 GetDefaultAudioSessionManager2(DataFlow dataFlow)
        {
            using (var enumerator = new MMDeviceEnumerator())
            {
                using (var device = enumerator.GetDefaultAudioEndpoint(dataFlow, Role.Multimedia))
                {
                    Debug.WriteLine("DefaultDevice: " + device.FriendlyName);
                    var sessionManager = AudioSessionManager2.FromMMDevice(device);
                    return sessionManager;
                }
            }
        }

        public void setVolume()
        {
            String processName = "TslGame";
            using (var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render)) {
                using (var sessionEnumerator = sessionManager.GetSessionEnumerator()) {
                    foreach (var session in sessionEnumerator) {
                        using (var control = session.QueryInterface<AudioSessionControl2>()) {
                            if (control.Process != null) {
                                //Console.WriteLine("Comparing: {0} ||| {1}",control.Process.ProcessName,processName);
                                if (control.Process.ProcessName.Contains(processName))
                                {
                                    using (var simpleVolume = session.QueryInterface<SimpleAudioVolume>())
                                    {
                                        bool muted = simpleVolume.IsMuted;
                                        simpleVolume.IsMuted = !muted;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /*public void petBattle() {
            while (isPetBattling)
            {
                String wowProcessName = "Wow";
                Random rnd = new Random();
                int lower = 3000;
                int upper = 6000;

                Process[] WowProcess = Process.GetProcessesByName(wowProcessName);
                if (WowProcess.Length > 0) {
                    System.Threading.Thread.Sleep(rnd.Next(lower, upper));
                    Console.WriteLine("!!! Pressing 'a'");
                    INPUT[] data = new INPUT[] {
                        new INPUT() {
                            type = INPUT_KEYBOARD,
                            u = new InputBatch
                            {
                                ki = new KEYBDINPUT
                                {
                                    wVk = 0x41,                        
                                    wScan = 0,    
                                    dwFlags = 0,
                                    dwExtraInfo = GetMessageExtraInfo(),
                                }
                            }
                        }
                    };
                    SendInput((uint)data.Length, data, Marshal.SizeOf(typeof(INPUT)));
                    //SendMessage(WowProcess[0].Handle, WM_KEYDOWN, 0x41, 0);
                    //SendMessage(WowProcess[0].Handle, WM_KEYUP, 0x41, 0);
                    //keybd_event(0x41, 0, KEYEVENTF_KEYDOWN, 0);
                    //keybd_event(0x41, 0, KEYEVENTF_KEYUP, 0);
                    //SendKeys.SendWait("a");
                }
            }
        }*/

        void runProgram(String filePath, String arguments)
        {
            Process proc = new Process();
            ProcessStartInfo pi = new ProcessStartInfo();

            pi.UseShellExecute = false;
            pi.FileName = @filePath;
            pi.Arguments = arguments;
            pi.CreateNoWindow = true;
            proc.StartInfo = pi;
            proc.Start();
        }
    }

    public sealed class KeyboardHook : IDisposable
    {
        // Registers a hot key with Windows.
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        // Unregisters the hot key with Windows.
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        /// <summary>
        /// Represents the window that is used internally to get the messages.
        /// </summary>
        private class Window : NativeWindow, IDisposable
        {
            private static int WM_HOTKEY = 0x0312;

            public Window()
            {
                // create the handle for the window.
                this.CreateHandle(new CreateParams());
            }

            /// <summary>
            /// Overridden to get the notifications.
            /// </summary>
            /// <param name="m"></param>
            protected override void WndProc(ref Message m)
            {
                base.WndProc(ref m);

                // check if we got a hot key pressed.
                if (m.Msg == WM_HOTKEY)
                {
                    // get the keys.
                    Keys key = (Keys)(((int)m.LParam >> 16) & 0xFFFF);
                    ModifierKeys modifier = (ModifierKeys)((int)m.LParam & 0xFFFF);

                    // invoke the event to notify the parent.
                    if (KeyPressed != null)
                        KeyPressed(this, new KeyPressedEventArgs(modifier, key));
                }
            }

            public event EventHandler<KeyPressedEventArgs> KeyPressed;

            #region IDisposable Members

            public void Dispose()
            {
                this.DestroyHandle();
            }

            #endregion
        }

        private Window _window = new Window();
        private int _currentId;

        public KeyboardHook()
        {
            // register the event of the inner native window.
            _window.KeyPressed += delegate (object sender, KeyPressedEventArgs args)
            {
                if (KeyPressed != null)
                    KeyPressed(this, args);
            };
        }

        /// <summary>
        /// Registers a hot key in the system.
        /// </summary>
        /// <param name="modifier">The modifiers that are associated with the hot key.</param>
        /// <param name="key">The key itself that is associated with the hot key.</param>
        public void RegisterHotKey(ModifierKeys modifier, Keys key)
        {
            // increment the counter.
            _currentId = _currentId + 1;

            // register the hot key.
            if (!RegisterHotKey(_window.Handle, _currentId, (uint)modifier, (uint)key))
                throw new InvalidOperationException("### Couldn’t register the hot key: " + modifier.ToString() + " | " + key.ToString());
        }

        /// <summary>
        /// A hot key has been pressed.
        /// </summary>
        public event EventHandler<KeyPressedEventArgs> KeyPressed;

        #region IDisposable Members

        public void Dispose()
        {
            // unregister all the registered hot keys.
            for (int i = _currentId; i > 0; i--)
            {
                UnregisterHotKey(_window.Handle, i);
            }

            // dispose the inner native window.
            _window.Dispose();
        }

        #endregion
    }

    /// <summary>
    /// Event Args for the event that is fired after the hot key has been pressed.
    /// </summary>
    public class KeyPressedEventArgs : EventArgs
    {
        private ModifierKeys _modifier;
        private Keys _key;

        internal KeyPressedEventArgs(ModifierKeys modifier, Keys key)
        {
            _modifier = modifier;
            _key = key;
        }

        public ModifierKeys Modifier
        {
            get { return _modifier; }
        }

        public Keys Key
        {
            get { return _key; }
        }
    }

    /*public struct INPUT
    {
        public int type;
        public InputBatch u;
    }*/

    [StructLayout(LayoutKind.Sequential)]
    public struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    /// <summary>
    /// The enumeration of possible modifiers.
    /// </summary>
    [Flags]
    public enum ModifierKeys : uint
    {
        Alt = 1,
        Control = 2,
        Shift = 4,
        Win = 8
    }
}

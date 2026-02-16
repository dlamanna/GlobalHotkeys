using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections.Concurrent;
using System.Text;
using NAudio.CoreAudioApi;

namespace GlobalHotkeys
{
    public partial class InvisibleForm : Form
    {
        private bool isPetBattling = false;
        public const int VK_MEDIA_PLAY_PAUSE = 0xB3;
        public const int VK_SPACE = 0x20;
        public const int KEYEVENTF_KEYDOWN = 0x0001;        //Key down flag
        public const int KEYEVENTF_KEYUP = 0x0002;          //Key up flag
        private readonly HashSet<string> currentAppsToMute = new(StringComparer.OrdinalIgnoreCase);
        private List<string> processNameList = [];
        private readonly HashSet<string> processNameSet = new(StringComparer.OrdinalIgnoreCase);
        private List<string> relevantSoundDevices = [];
        //private readonly string nirCMDPath = @"C:\Windows\SysWOW64\nircmd.exe";
        private readonly string petBattlePath = @"C:\Program Files (x86)\AutoIt3\USERCREATED\simplePetBattle.exe";
        private readonly string tarkovFleaRefreshPath = @"C:\Program Files (x86)\AutoIt3\USERCREATED\tarkovFleaRefresh.exe";
        private readonly string crosshairPath = @"C:\Program Files (x86)\DesktopShell\Bin\Crosshair.exe";
        private readonly string tooltipPath = @"C:\Program Files\DesktopShell\Bin\ToolTipper.exe";
        private readonly KeyboardHook hook = new KeyboardHook();

        // Performance optimization: Cache for process lookups
        private readonly Dictionary<string, int> processCache = new(StringComparer.OrdinalIgnoreCase);
        private System.Threading.Timer? processCacheRefreshTimer;
        private readonly object processCacheLock = new();

        // Performance optimization: Audio session cache
        private readonly Dictionary<string, List<AudioSessionInfo>> audioSessionCache = [];
        private System.Threading.Timer? audioCacheRefreshTimer;
        private readonly object audioCacheLock = new();

        // Performance optimization: Background thread for logging
        private readonly ConcurrentQueue<string> logQueue = new();
        private readonly AutoResetEvent logSignal = new(false);
        private Thread? logThread;
        private volatile bool logThreadRunning = true;
        private StreamWriter? logWriter;
        private readonly object logWriterLock = new();
        private static readonly ConcurrentDictionary<Type, System.Reflection.MemberInfo?> SessionPidMemberCache = new();

        // Audio session info class for caching
        private class AudioSessionInfo
        {
            public string ProcessName { get; set; } = string.Empty;
            public AudioSessionControl Session { get; set; } = null!;
            public string DeviceName { get; set; } = string.Empty;
            public bool IsMuted { get; set; } = false;
        }

        private static int? TryGetSessionProcessIdFast(AudioSessionControl session)
        {
            var type = session.GetType();

            var member = SessionPidMemberCache.GetOrAdd(type, static t =>
                (System.Reflection.MemberInfo?)t.GetProperty("GetProcessID")
                ?? t.GetMethod("GetProcessID", Type.EmptyTypes));

            try
            {
                if (member is System.Reflection.PropertyInfo prop)
                {
                    var value = prop.GetValue(session);
                    return value switch
                    {
                        int i when i > 0 => i,
                        uint u when u > 0 => unchecked((int)u),
                        _ => null
                    };
                }

                if (member is System.Reflection.MethodInfo method)
                {
                    var value = method.Invoke(session, null);
                    return value switch
                    {
                        int i when i > 0 => i,
                        uint u when u > 0 => unchecked((int)u),
                        _ => null
                    };
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        // Method to get the actual process name from an audio session
        private string? GetProcessNameFromSession(AudioSessionControl session)
        {
            try
            {
                // Fast path: PID -> process name (cached reflection)
                int? pid = TryGetSessionProcessIdFast(session);
                if (pid is int p && p > 0)
                {
                    try
                    {
                        using var process = Process.GetProcessById(p);
                        if (!string.IsNullOrEmpty(process.ProcessName))
                        {
                            return process.ProcessName;
                        }
                    }
                    catch
                    {
                        // Process might not exist anymore, continue
                    }
                }

                // First, try to access the GetProcessID property directly
                var sessionType = session.GetType();
                var processIdProperty = sessionType.GetProperty("GetProcessID");
                
                if (processIdProperty != null)
                {
                    try
                    {
                        var processId = processIdProperty.GetValue(session);
                        if (processId != null)
                        {
                            if (processId is uint processIdUint && processIdUint > 0)
                            {
                                try
                                {
                                    using var process = Process.GetProcessById((int)processIdUint);
                                    if (!string.IsNullOrEmpty(process.ProcessName))
                                    {
                                        //Log($"%%% Found process: {process.ProcessName} (PID: {processIdUint})");
                                        return process.ProcessName;
                                    }
                                }
                                catch
                                {
                                    // Process might not exist anymore, continue
                                }
                            }
                            else if (processId is int processIdInt && processIdInt > 0)
                            {
                                try
                                {
                                    using var process = Process.GetProcessById(processIdInt);
                                    if (!string.IsNullOrEmpty(process.ProcessName))
                                    {
                                        //Log($"%%% Found process: {process.ProcessName} (PID: {processIdInt})");
                                        return process.ProcessName;
                                    }
                                }
                                catch
                                {
                                    // Process might not exist anymore, continue
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        //Log("%%% Error accessing GetProcessID property");
                    }
                }
                
                // If direct property access fails, try reflection on internal fields
                var fields = sessionType.GetFields(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                
                foreach (var field in fields)
                {
                    try
                    {
                        var value = field.GetValue(session);
                        if (value != null)
                        {
                            var valueType = value.GetType();
                            
                            // Look for methods that can give us the process ID
                            var methods = valueType.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                            
                            foreach (var method in methods)
                            {
                                try
                                {
                                    if (method.Name.Contains("Process") || method.Name.Contains("ProcessId") || method.Name.Contains("GetProcess"))
                                    {
                                        var result = method.Invoke(value, null);
                                        if (result != null)
                                        {
                                            if (result is uint processId)
                                            {
                                                var actualProcessId = processId;
                                                if (actualProcessId > 0)
                                                {
                                                    try
                                                    {
                                                        using var process = Process.GetProcessById((int)actualProcessId);
                                                        if (!string.IsNullOrEmpty(process.ProcessName))
                                                        {
                                                            //Log($"%%% Found process: {process.ProcessName} (PID: {actualProcessId})");
                                                            return process.ProcessName;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        // Process might not exist anymore, continue
                                                    }
                                                }
                                            }
                                            else if (result is int processIdInt)
                                            {
                                                var actualProcessId = (uint)processIdInt;
                                                if (actualProcessId > 0)
                                                {
                                                    try
                                                    {
                                                        using var process = Process.GetProcessById(processIdInt);
                                                        if (!string.IsNullOrEmpty(process.ProcessName))
                                                        {
                                                            //Log($"%%% Found process: {process.ProcessName} (PID: {actualProcessId})");
                                                            return process.ProcessName;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        // Process might not exist anymore, continue
                                                    }
                                                }
                                            }
                                            else if (result is string processName && !string.IsNullOrEmpty(processName))
                                            {
                                                //Log($"%%% Found process name directly: {processName}");
                                                return processName;
                                            }
                                        }
                                    }
                                }
                                catch
                                {
                                    // Continue to next method
                                }
                            }
                            
                            // Also check properties
                            var properties = valueType.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                            foreach (var prop in properties)
                            {
                                try
                                {
                                    if (prop.Name.Contains("Process") || prop.Name.Contains("ProcessId"))
                                    {
                                        var result = prop.GetValue(value);
                                        if (result != null)
                                        {
                                            if (result is uint processId)
                                            {
                                                var actualProcessId = processId;
                                                if (actualProcessId > 0)
                                                {
                                                    try
                                                    {
                                                        using var process = Process.GetProcessById((int)actualProcessId);
                                                        if (!string.IsNullOrEmpty(process.ProcessName))
                                                        {
                                                            //Log($"%%% Found process: {process.ProcessName} (PID: {actualProcessId})");
                                                            return process.ProcessName;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        // Process might not exist anymore, continue
                                                    }
                                                }
                                            }
                                            else if (result is int processIdInt)
                                            {
                                                var actualProcessId = (uint)processIdInt;
                                                if (actualProcessId > 0)
                                                {
                                                    try
                                                    {
                                                        using var process = Process.GetProcessById(processIdInt);
                                                        if (!string.IsNullOrEmpty(process.ProcessName))
                                                        {
                                                            //Log($"%%% Found process: {process.ProcessName} (PID: {actualProcessId})");
                                                            return process.ProcessName;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        // Process might not exist anymore, continue
                                                    }
                                                }
                                            }
                                            else if (result is string processName && !string.IsNullOrEmpty(processName))
                                            {
                                                //Log($"%%% Found process: {processName}");
                                                return processName;
                                            }
                                        }
                                    }
                                }
                                catch
                                {
                                    // Continue to next property
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Continue to next field
                    }
                }
                
                // If we still can't find it, try a different approach - look for the session identifier
                try
                {
                    var identifierMethod = sessionType.GetMethod("GetSessionIdentifier", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                    if (identifierMethod != null)
                    {
                        var identifier = identifierMethod.Invoke(session, null) as string;
                        if (!string.IsNullOrEmpty(identifier))
                        {
                            if (identifier.Contains('\\'))
                            {
                                var parts = identifier.Split('\\');
                                foreach (var part in parts)
                                {
                                    if (!string.IsNullOrEmpty(part) && !part.Contains('.') && part.Length > 2)
                                    {
                                        return part;
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // Continue
                }
            }
            catch (Exception ex)
            {
                Log($"### Error getting process name from session: {ex.Message}");
            }
            
            return null;
        }

        // NAudio audio device enumerator
        private MMDeviceEnumerator? audioDeviceEnumerator;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);



        public InvisibleForm()
        {
            InitializeComponent();

            ResetLog();

            // Initialize performance optimizations
            InitializePerformanceOptimizations();

            // register the event that is fired after the key press.
            hook.KeyPressed += new EventHandler<KeyPressedEventArgs>(Hook_KeyPressed);

            GlobalHotkeys.ModifierKeys ctrl = GlobalHotkeys.ModifierKeys.Control;
            // Mute/Unmute Game Toggle (Defaults: PUBG/Tarkov/Naraka/NewWorld/Battlefield/CoD/WoW/LostArk/BDO)
            _ = IsHotkeyRegistrationSuccessful(ctrl, Keys.NumPad1);

            // Try to add current window to process toggle list
            _ = IsHotkeyRegistrationSuccessful(ctrl, Keys.NumPad3);

            // Toggle Pet Battle (semicolon-key spam) or Tarkov flea market spam (F5)
            _ = IsHotkeyRegistrationSuccessful(ctrl, Keys.NumPad4);

            // Toggle Crosshair program
            _ = IsHotkeyRegistrationSuccessful(ctrl, Keys.Multiply);

            // Initialize audio device enumeration
            try
            {
                audioDeviceEnumerator = new MMDeviceEnumerator();
                Log("%%% Audio system initialized");
            }
            catch (Exception ex)
            {
                Log($"### Failed to initialize audio device enumerator: {ex.Message}");
                audioDeviceEnumerator = null;
            }

            // Get all relevant audio devices from SoundDevices.txt
            ScanSoundDevices();

            // Get all applications we want to be mutable from ProcessList.txt
            ScanApplications();

            // Initial cache population
            Task.Run(() => {
                RefreshProcessCache();
                RefreshAudioSessionCache();
            });
        }

        private void InitializePerformanceOptimizations()
        {
            // Initialize process cache refresh timer (every 5 seconds)
            processCacheRefreshTimer = new System.Threading.Timer(RefreshProcessCache, null, 5000, 5000);

            // Initialize audio session cache refresh timer (every 10 seconds)
            audioCacheRefreshTimer = new System.Threading.Timer(RefreshAudioSessionCache, null, 10000, 10000);

            // Initialize background logging thread
            logThread = new Thread(ProcessLogQueue)
            {
                IsBackground = true,
                Priority = ThreadPriority.BelowNormal
            };
            logThread?.Start();
        }

        private void RefreshProcessCache(object? state = null)
        {
            try
            {
                if (processNameSet.Count == 0)
                {
                    lock (processCacheLock)
                    {
                        processCache.Clear();
                    }
                    return;
                }

                // IMPORTANT: do not cache Process objects (they hold OS handles).
                // Cache only counts to keep the hotkey path fast without leaking handles.
                var processCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var process in Process.GetProcesses())
                {
                    try
                    {
                        var name = process.ProcessName;
                        if (processNameSet.Contains(name))
                        {
                            processCounts.TryGetValue(name, out int count);
                            processCounts[name] = count + 1;
                        }
                    }
                    catch
                    {
                        // Ignore individual process errors
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
                
                lock (processCacheLock)
                {
                    processCache.Clear();
                    foreach (var kvp in processCounts)
                    {
                        processCache[kvp.Key] = kvp.Value;
                    }
                }
            }
            catch (Exception) { /* Ignore cache refresh errors */ }
        }

        private void RefreshAudioSessionCache(object? state = null)
        {
            // Early exit if no processes to monitor
            if (processNameSet.Count == 0)
            {
                Dictionary<string, List<AudioSessionInfo>>? oldCache = null;
                lock (audioCacheLock)
                {
                    oldCache = new Dictionary<string, List<AudioSessionInfo>>(audioSessionCache, StringComparer.OrdinalIgnoreCase);
                    audioSessionCache.Clear();
                }

                if (oldCache != null)
                {
                    foreach (var kvp in oldCache)
                    foreach (var sessionInfo in kvp.Value)
                        try { sessionInfo.Session?.Dispose(); } catch { }
                }

                return;
            }
            try
            {
                if (audioDeviceEnumerator == null) return;

                string? perfVar = Environment.GetEnvironmentVariable("GLOBALHOTKEYS_PERF");
                bool perfEnabled = !string.IsNullOrWhiteSpace(perfVar)
                    && !string.Equals(perfVar.Trim(), "0", StringComparison.Ordinal)
                    && !string.Equals(perfVar.Trim(), "false", StringComparison.OrdinalIgnoreCase);
                Stopwatch? sw = perfEnabled ? Stopwatch.StartNew() : null;

                // Per-refresh PID->processName cache to avoid repeated Process.GetProcessById calls.
                var pidToName = new Dictionary<int, string>();

                int deviceCount = 0;
                int relevantDeviceCount = 0;
                int sessionCount = 0;
                int keptSessions = 0;
                int pidHits = 0;
                int pidMisses = 0;

                var newAudioSessionCache = new Dictionary<string, List<AudioSessionInfo>>(StringComparer.OrdinalIgnoreCase);
                
                // Enumerate all audio devices
                var devices = audioDeviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
                try
                {
                    foreach (var device in devices)
                    {
                        try
                        {
                            deviceCount++;

                            // Check if this device is in our list of relevant devices
                            bool isRelevantDevice = relevantSoundDevices.Count == 0;
                            if (!isRelevantDevice)
                            {
                                string friendly = device.FriendlyName;
                                for (int d = 0; d < relevantSoundDevices.Count; d++)
                                {
                                    if (friendly.Contains(relevantSoundDevices[d], StringComparison.OrdinalIgnoreCase))
                                    {
                                        isRelevantDevice = true;
                                        break;
                                    }
                                }
                            }

                            if (!isRelevantDevice)
                            {
                                continue;
                            }

                            relevantDeviceCount++;

                            var sessions = device.AudioSessionManager?.Sessions;
                            if (sessions != null)
                            {
                                for (int i = 0; i < sessions.Count; i++)
                                {
                                    var session = sessions[i];
                                    sessionCount++;

                                    // Try to get the actual process name from the session
                                    string? processName = null;

                                    try
                                    {
                                        // Fast path: PID -> process name (avoid reflection heavy fallbacks and repeated Process.GetProcessById)
                                        int? pid = TryGetSessionProcessIdFast(session);
                                        if (pid is int p && p > 0)
                                        {
                                            if (pidToName.TryGetValue(p, out var cachedName))
                                            {
                                                processName = cachedName;
                                                if (perfEnabled) pidHits++;
                                            }
                                            else
                                            {
                                                if (perfEnabled) pidMisses++;
                                                try
                                                {
                                                    using var proc = Process.GetProcessById(p);
                                                    if (!string.IsNullOrEmpty(proc.ProcessName))
                                                    {
                                                        pidToName[p] = proc.ProcessName;
                                                        processName = proc.ProcessName;
                                                    }
                                                }
                                                catch
                                                {
                                                    // Process may have exited
                                                }
                                            }
                                        }

                                        processName ??= GetProcessNameFromSession(session);

                                        if (string.IsNullOrEmpty(processName))
                                        {
                                            // Fallback to generic name if we can't get the process name
                                            processName = $"Session_{i}";
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"### Error getting session info for session {i}: {ex.Message}");
                                        processName = $"Session_{i}";
                                    }

                                    // Fast path: only cache sessions for relevant processes
                                    if (!processNameSet.Contains(processName))
                                    {
                                        session.Dispose();
                                        continue;
                                    }

                                    keptSessions++;

                                    if (!newAudioSessionCache.TryGetValue(processName, out var list))
                                    {
                                        list = [];
                                        newAudioSessionCache[processName] = list;
                                    }

                                    list.Add(new AudioSessionInfo
                                    {
                                        ProcessName = processName,
                                        Session = session,
                                        DeviceName = device.FriendlyName,
                                        IsMuted = session.SimpleAudioVolume?.Mute ?? false
                                    });
                                }
                            }
                        }
                        catch (Exception)
                        {
                            /* Ignore individual device errors */
                        }
                        finally
                        {
                            device?.Dispose();
                        }
                    }
                }
                finally
                {
                    (devices as IDisposable)?.Dispose();
                }
                
                Dictionary<string, List<AudioSessionInfo>>? oldCache = null;
                lock (audioCacheLock)
                {
                    oldCache = new Dictionary<string, List<AudioSessionInfo>>(audioSessionCache, StringComparer.OrdinalIgnoreCase);
                    audioSessionCache.Clear();
                    foreach (var kvp in newAudioSessionCache)
                    {
                        audioSessionCache[kvp.Key] = kvp.Value;
                    }
                }

                // Dispose old COM objects outside lock
                if (oldCache != null)
                {
                    foreach (var kvp in oldCache)
                    {
                        foreach (var sessionInfo in kvp.Value)
                        {
                            try { sessionInfo.Session?.Dispose(); } catch { }
                        }
                    }
                }
                
                // Only log cache status if there are relevant processes found
                if (newAudioSessionCache.Count > 0)
                {
                    //Log($"%%% Audio cache updated: {newAudioSessionCache.Count} relevant processes found");
                }

                if (perfEnabled && sw != null)
                {
                    sw.Stop();
                    Log($"PERF AudioCache: {sw.ElapsedMilliseconds}ms devices={deviceCount} relevantDevices={relevantDeviceCount} sessions={sessionCount} kept={keptSessions} pidHits={pidHits} pidMisses={pidMisses} pidCache={pidToName.Count}");
                }
            }
            catch (Exception ex)
            {
                Log($"### Error refreshing audio session cache: {ex.Message}");
            }
        }

        private void ProcessLogQueue()
        {
            try
            {
                var logPath = Path.Combine(AppContext.BaseDirectory, "GlobalHotkeys.log");
                lock (logWriterLock)
                {
                    logWriter = new StreamWriter(new FileStream(logPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite), Encoding.UTF8)
                    {
                        AutoFlush = true
                    };
                }

                while (logThreadRunning)
                {
                    // Sleep efficiently until we have work (or periodically wake to allow shutdown)
                    logSignal.WaitOne(250);

                    while (logQueue.TryDequeue(out var logEntry))
                    {
                        try
                        {
                            lock (logWriterLock)
                            {
                                logWriter?.WriteLine(logEntry);
                            }
                        }
                        catch
                        {
                            // Ignore logging errors
                        }
                    }
                }
            }
            catch
            {
                // Ignore logging init errors
            }
            finally
            {
                lock (logWriterLock)
                {
                    logWriter?.Dispose();
                    logWriter = null;
                }
            }
        }

        private void DisposeResources()
        {
            try
            {
                logThreadRunning = false;
                logSignal.Set();

                processCacheRefreshTimer?.Dispose();
                audioCacheRefreshTimer?.Dispose();
                audioDeviceEnumerator?.Dispose();
                hook.Dispose();

                if (logThread != null && logThread.IsAlive)
                {
                    // Avoid hanging on exit
                    logThread.Join(500);
                }
            }
            catch
            {
                // Best-effort cleanup
            }
        }

        private bool IsHotkeyRegistrationSuccessful(ModifierKeys keys, Keys key)
        {
            bool isSuccess = hook.RegisterHotKey(keys, key);
            if (!isSuccess) {
                ToolTip("GlobalHotkey Error", $"### IsHotkeyRegistrationSuccessful - Error binding ({keys}){key}");
            }
            else {
                Log($"$$$ Successfully bound ({keys}){key}");
            }
            return isSuccess;
        }

        public void ToolTip(string title, string message) 
        {
            RunProgram(tooltipPath, $"\"{title}\"", $"\"{message}\"");
        }

        public void ScanApplications()
        {
            string applicationListPath = "ProcessList.txt";
            if (!File.Exists(applicationListPath))
            {
                processNameList = [];
                processNameSet.Clear();
                Log($"### Missing {applicationListPath}; no processes will be toggled");
                return;
            }

            processNameList = [.. File.ReadLines(applicationListPath)
                .Select(line => line.Trim())
                .Where(line => line.Length > 0 && !line.StartsWith('#'))
                .Distinct(StringComparer.OrdinalIgnoreCase)];

            processNameSet.Clear();
            foreach (var name in processNameList)
            {
                processNameSet.Add(name);
            }

            foreach (string process in processNameList)
            {
                Log($"%%% Process scanned in: {process}");
            }
        }

        public void SaveApplications()
        {
            string applicationListPath = "ProcessList.txt";
            File.WriteAllLines(applicationListPath, processNameList);
            Log($"%%% Processes saved to: {applicationListPath}");
        }

        public void ScanSoundDevices()
        {
            string soundDevicesPath = "SoundDevices.txt";
            if (!File.Exists(soundDevicesPath))
            {
                relevantSoundDevices = [];
                Log($"### Missing {soundDevicesPath}; all devices will be ignored");
                return;
            }

            relevantSoundDevices = [.. File.ReadLines(soundDevicesPath)
                .Select(line => line.Trim())
                .Where(line => line.Length > 0 && !line.StartsWith('#'))
                .Distinct(StringComparer.OrdinalIgnoreCase)];

            foreach (string device in relevantSoundDevices)
            {
                Log($"$$$ Sound device scanned in: {device}");
            }
        }

        public void Log(params string[] logOutput)
        {
            string combinedLogOutput = $"{DateTime.Now:HH:mm:ss.fff}:\t{string.Join(" ", logOutput)}";
            
            // Queue log entry for background processing to avoid blocking
            logQueue.Enqueue(combinedLogOutput);
            logSignal.Set();
            
            // Still output to console immediately for debugging
            Console.WriteLine(combinedLogOutput);
        }
        
        public void ResetLog()
        {
            try
            {
                using (File.Create("GlobalHotkeys.log")) { };
            }
            catch (Exception e)
            {
                ToolTip("Error", $"ResetLog() - {e.GetType()}\n{e.Message}");
            }
        }

        private void Hook_KeyPressed(object? sender, KeyPressedEventArgs e)
        {
            Log($"^^^ Modifier: {e.Modifier}({(int)e.Modifier}),\t Key: {e.Key}({(int)e.Key})");

            switch((int)e.Key) {
                case 19:
                    Log("!!! Do Firefox Toggle");
                    Process[] FirefoxProcesses = Process.GetProcessesByName("Firefox");
                    Process? FirefoxProcess = null;
                    foreach (Process p in FirefoxProcesses)
                    {
                        //Log($"Process: {p.Id}\t{p.MainWindowTitle}");
                        if (p.MainWindowTitle.Length > 0)
                        {
                            FirefoxProcess = p;
                            Log($"@@@ Found active tab: {FirefoxProcess.MainWindowTitle}");
                            break;
                        }
                    }
                    
                    if (FirefoxProcess != null && FirefoxProcess.MainWindowHandle != IntPtr.Zero)
                    {
                        var procID = FirefoxProcess.MainWindowHandle;
                        Keyboard.Key firefoxHotkey = new Keyboard.Key(Keyboard.Messaging.VKeys.KEY_P, Keyboard.Messaging.VKeys.KEY_LSHIFT, Keyboard.Messaging.ShiftType.SHIFT_ALT); //ALT + SHIFT + P
                        firefoxHotkey.PressBackground(procID);
                    }
                    else
                    {
                        Log("### No Firefox process with active window found");
                    }
                    break;

                case 97:
                    Log("!!! Do Game Mute Toggle");
                    currentAppsToMute.Clear();
                    
                    // Use cached process lookups for much faster performance
                    lock (processCacheLock)
                    {
                        foreach (string processNameToLookFor in processNameList)
                        {
                            if (processCache.TryGetValue(processNameToLookFor, out int count) && count > 0)
                            {
                                currentAppsToMute.Add(processNameToLookFor);
                                Log($"\t@@@ Game Found: '{processNameToLookFor}' (cached x{count})");
                            }
                            else
                            {
                                Log($"\t!!! No '{processNameToLookFor}' process found");
                            }
                        }
                    }

                    if (currentAppsToMute.Count > 0)
                    {
                        // Execute mute immediately for maximum responsiveness
                        SetVolume();
                        // Refresh audio cache in background for future use
                        Task.Run(() => RefreshAudioSessionCache());
                    }

                    break;

                case 99:
                    var foregroundProcess = GetForegroundProcess();
                    if (foregroundProcess == null)
                    {
                        Log("### Could not get foreground process");
                        break;
                    }
                    
                    string procNameTemp = foregroundProcess.ProcessName;
                    Log($"@@@ Foreground Process Name: '{procNameTemp}'");
                    if (processNameSet.Contains(procNameTemp)) 
                    {
                        Log($"### Process: '{procNameTemp}' already exists in mute list, skipping");
                    }
                    else 
                    {
                        processNameList.Add(procNameTemp);
                        processNameSet.Add(procNameTemp);
                        Log($"$$$ Adding process: '{procNameTemp}' to mute list and saving to file");
                        SaveApplications();
                        RefreshProcessCache(); // Refresh cache after adding new process
                    }
                    break;                

                case 100:
                    Log("!!! Do PetBattle/Tarkov Toggle");
                    isPetBattling = !isPetBattling;
                    if(isPetBattling) 
                    {
                        Process[] prunedProcessList = Process.GetProcessesByName("EscapeFromTarkov");
                        if(prunedProcessList.Length >= 1) {
                            RunProgram(tarkovFleaRefreshPath);
                        }
                        else {
                            RunProgram(petBattlePath);
                        }
                    }
                    else {
                        KillProgram("simplepetbattle_tarkovflearefresh");
                    }
                    break;

                case 106:
                    Log("!!! Do Crosshair Toggle");
                    string exeName = "Crosshair";
                    bool isRunning = false;
                    Process[] processList = Process.GetProcessesByName(exeName);
                    if(processList.Length >= 1) 
                    {
                        foreach(Process p in processList) 
                        {
                            try
                            {
                                Log($"@@@ Process Found: {p.ProcessName}\t{{{p.Id}}}");
                                p.Kill();
                            }
                            catch (Exception ex)
                            {
                                Log($"### Error killing process: {ex.Message}");
                            }
                        }
                        isRunning = true;
                    }

                    if(!isRunning) {
                        RunProgram(crosshairPath, "");
                    }
                    break;
                default:
                    break;
            }
        }



                public void SetVolume()
        {
            try
            {
                if (audioDeviceEnumerator == null)
                {
                    Log("### Audio device enumerator not available");
                    return;
                }

                Log($"@@@ Attempting to mute {currentAppsToMute.Count} processes");
                
                // Only mute the specific processes in our list
                lock (audioCacheLock)
                {
                    if (audioSessionCache.Count > 0)
                    {
                        Log($"@@@ Found {audioSessionCache.Count} cached audio sessions to process");
                        
                        foreach (var kvp in audioSessionCache)
                        {
                            string processName = kvp.Key;
                            var sessions = kvp.Value;
                            
                            // Only process sessions for processes in our mute list
                            if (currentAppsToMute.Contains(processName, StringComparer.OrdinalIgnoreCase))
                            {
                                Log($"@@@ Processing {sessions.Count} sessions for '{processName}' (in mute list)");
                                
                                foreach (var sessionInfo in sessions)
                                {
                                    try
                                    {
                                        var simpleVolume = sessionInfo.Session.SimpleAudioVolume;
                                        if (simpleVolume != null)
                                        {
                                            // Toggle mute state
                                            bool newMuteState = !simpleVolume.Mute;
                                            simpleVolume.Mute = newMuteState;
                                            
                                            Log($"@@@ Toggled mute on {processName} ({sessionInfo.DeviceName}): {(newMuteState ? "MUTED" : "UNMUTED")}");
                                            
                                            // Update cache
                                            sessionInfo.IsMuted = newMuteState;
                                        }
                                        else
                                        {
                                            Log($"### SimpleAudioVolume is null for session {processName}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"### Error toggling mute for {processName} on {sessionInfo.DeviceName}: {ex.Message}");
                                    }
                                }
                            }
                            else
                            {
                                Log($"@@@ Skipping {processName} (not in current mute list)");
                            }
                        }
                    }
                    else
                    {
                        Log($"### No cached audio sessions found. Audio cache count: {audioSessionCache.Count}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"### SetVolume() - {ex.Message}");
            }
            finally
            {
                currentAppsToMute.Clear();
            }
        }

        private void MuteProcessDirectly(string processName)
        {
            try
            {
                var devices = audioDeviceEnumerator?.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
                if (devices == null) return;

                foreach (var device in devices)
                {
                    try
                    {
                        // Check if this device is in our list of relevant devices
                        bool isRelevantDevice = false;
                        if (relevantSoundDevices.Count > 0)
                        {
                            string friendly = device.FriendlyName;
                            for (int d = 0; d < relevantSoundDevices.Count; d++)
                            {
                                if (friendly.Contains(relevantSoundDevices[d], StringComparison.OrdinalIgnoreCase))
                                {
                                    isRelevantDevice = true;
                                    break;
                                }
                            }
                        }
                        
                        if (!isRelevantDevice)
                        {
                            continue; // Skip non-relevant devices
                        }
                        
                        var sessions = device.AudioSessionManager?.Sessions;
                        if (sessions != null)
                        {
                            for (int i = 0; i < sessions.Count; i++)
                            {
                                var session = sessions[i];
                                
                                // Try to get the actual process name from the session
                                string? sessionProcessName = null;
                                
                                try
                                {
                                    // Use reflection to access internal properties if available
                                    var processProperty = session.GetType().GetProperty("Process");
                                    if (processProperty != null)
                                    {
                                        var process = processProperty.GetValue(session);
                                        if (process != null)
                                        {
                                            var processNameProperty = process.GetType().GetProperty("ProcessName");
                                            if (processNameProperty != null)
                                            {
                                                sessionProcessName = processNameProperty.GetValue(process) as string;
                                            }
                                        }
                                    }
                                }
                                catch (Exception) { /* Ignore reflection errors */ }
                                
                                // Only mute if this session belongs to the target process
                                if (!string.IsNullOrEmpty(sessionProcessName) && 
                                    sessionProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase))
                                {
                                    var simpleVolume = session.SimpleAudioVolume;
                                    if (simpleVolume != null)
                                    {
                                        bool newMuteState = !simpleVolume.Mute;
                                        simpleVolume.Mute = newMuteState;
                                        
                                        Log($"@@@ Direct mute toggle on {processName} ({device.FriendlyName}): {(newMuteState ? "MUTED" : "UNMUTED")}");
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception) { /* Ignore individual device errors */ }
                    finally
                    {
                        device?.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"### Error in direct mute for {processName}: {ex.Message}");
            }
        }

        public static Process? GetForegroundProcess()
        {
            IntPtr hWnd = GetForegroundWindow();                                    // Get foreground window handle
            if (GetWindowThreadProcessId(hWnd, out uint processID) == 0)           // Check if we got a valid PID
            {
                return null;
            }
            
            try
            {
                Process fgProc = Process.GetProcessById(Convert.ToInt32(processID));    // Get it as a C# obj.
                return fgProc;
            }
            catch (ArgumentException)
            {
                // Process might have terminated
                return null;
            }
        }

        private IntPtr RunProgram(string filePath, params string[] arguments)
        {
            using var proc = new Process();

            proc.StartInfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                FileName = filePath,
                Arguments = string.Join(" ", arguments),
                CreateNoWindow = true,
                // Optional clarity: when UseShellExecute=false, this is respected for console apps
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = Path.GetDirectoryName(filePath) ?? AppContext.BaseDirectory,
            };

            proc.Start();

            // Note: for GUI apps this can be 0 immediately; current code already assumes that.
            return proc.MainWindowHandle;
        }

        private void KillProgram(string exeName)
        {
            string[] exeNames = exeName.Split('_');
            foreach(string s in exeNames) 
            {
                try
                {
                    Process[] processList = Process.GetProcessesByName(s);
                    if(processList.Length >= 1) 
                    {
                        foreach(Process p in processList) 
                        {
                            try
                            {
                                Console.WriteLine($"Process Found: {p.ProcessName}\t{{{p.Id}}}");
                                p.Kill();
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine($"Error killing process {p.ProcessName}: {e.Message}");
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error finding processes for {s}: {e.Message}");
                }
            }
        }

    }
}
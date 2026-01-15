using System.Diagnostics;
using System.Runtime.InteropServices;
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
        private List<string> currentAppsToMute = [];
        private List<string> processNameList = [];
        private List<string> relevantSoundDevices = [];
        //private readonly string nirCMDPath = @"C:\Windows\SysWOW64\nircmd.exe";
        private readonly string petBattlePath = @"C:\Program Files (x86)\AutoIt3\USERCREATED\simplePetBattle.exe";
        private readonly string tarkovFleaRefreshPath = @"C:\Program Files (x86)\AutoIt3\USERCREATED\tarkovFleaRefresh.exe";
        private readonly string crosshairPath = @"C:\Program Files (x86)\DesktopShell\Bin\Crosshair.exe";
        private readonly string tooltipPath = @"C:\Program Files\DesktopShell\Bin\ToolTipper.exe";
        private readonly KeyboardHook hook = new KeyboardHook();

        // Performance optimization: Cache for process lookups
        private readonly Dictionary<string, Process[]> processCache = [];
        private System.Threading.Timer? processCacheRefreshTimer;
        private readonly object processCacheLock = new();

        // Performance optimization: Audio session cache
        private readonly Dictionary<string, List<AudioSessionInfo>> audioSessionCache = [];
        private System.Threading.Timer? audioCacheRefreshTimer;
        private readonly object audioCacheLock = new();

        // Performance optimization: Background thread for logging
        private readonly Queue<string> logQueue = new();
        private readonly object logLock = new();
        private Thread? logThread;
        private volatile bool logThreadRunning = true;

        // Audio session info class for caching
        private class AudioSessionInfo
        {
            public string ProcessName { get; set; } = string.Empty;
            public AudioSessionControl Session { get; set; } = null!;
            public string DeviceName { get; set; } = string.Empty;
            public bool IsMuted { get; set; } = false;
        }

        // Method to get the actual process name from an audio session
        private string? GetProcessNameFromSession(AudioSessionControl session)
        {
            try
            {
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
                                    var process = Process.GetProcessById((int)processIdUint);
                                    if (process != null && !string.IsNullOrEmpty(process.ProcessName))
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
                                    var process = Process.GetProcessById(processIdInt);
                                        if (process != null && !string.IsNullOrEmpty(process.ProcessName))
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
                                                        var process = Process.GetProcessById((int)actualProcessId);
                                                        if (process != null && !string.IsNullOrEmpty(process.ProcessName))
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
                                                        var process = Process.GetProcessById(processIdInt);
                                                        if (process != null && !string.IsNullOrEmpty(process.ProcessName))
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
                                                        var process = Process.GetProcessById((int)actualProcessId);
                                                        if (process != null && !string.IsNullOrEmpty(process.ProcessName))
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
                                                        var process = Process.GetProcessById(processIdInt);
                                                        if (process != null && !string.IsNullOrEmpty(process.ProcessName))
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

            // Initialize performance optimizations
            InitializePerformanceOptimizations();

            // register the event that is fired after the key press.
            hook.KeyPressed += new EventHandler<KeyPressedEventArgs>(Hook_KeyPressed);
            ResetLog();

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
                // Use a more efficient approach - get all processes once and filter
                var allProcesses = Process.GetProcesses();
                var processDict = new Dictionary<string, Process[]>(StringComparer.OrdinalIgnoreCase);
                
                foreach (string processName in processNameList)
                {
                    try
                    {
                        var matchingProcesses = allProcesses
                            .Where(p => string.Equals(p.ProcessName, processName, StringComparison.OrdinalIgnoreCase))
                            .ToArray();
                        
                        if (matchingProcesses.Length > 0)
                        {
                            processDict[processName] = matchingProcesses;
                        }
                    }
                    catch (Exception) { /* Ignore individual process errors */ }
                }
                
                lock (processCacheLock)
                {
                    processCache.Clear();
                    foreach (var kvp in processDict)
                    {
                        processCache[kvp.Key] = kvp.Value;
                    }
                }
            }
            catch (Exception) { /* Ignore cache refresh errors */ }
        }

        private void RefreshAudioSessionCache(object? state = null)
        {
            try
            {
                if (audioDeviceEnumerator == null) return;

                var newAudioSessionCache = new Dictionary<string, List<AudioSessionInfo>>(StringComparer.OrdinalIgnoreCase);
                
                // Enumerate all audio devices
                var devices = audioDeviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
                
                foreach (var device in devices)
                {
                    try
                    {
                        // Check if this device is in our list of relevant devices
                        bool isRelevantDevice = relevantSoundDevices.Any(deviceName => 
                            device.FriendlyName.Contains(deviceName, StringComparison.OrdinalIgnoreCase));
                        
                        if (!isRelevantDevice)
                        {
                            //Log($"%%% Skipping audio device: {device.FriendlyName} (not in SoundDevices.txt)");
                            continue;
                        }
                        
                        var sessions = device.AudioSessionManager?.Sessions;
                        if (sessions != null)
                        {
                            for (int i = 0; i < sessions.Count; i++)
                            {
                                var session = sessions[i];
                                
                                // Try to get the actual process name from the session
                                string? processName = null;
                                
                                try
                                {
                                    // Use our new method to get the actual process name
                                    processName = GetProcessNameFromSession(session);
                                    
                                    if (string.IsNullOrEmpty(processName))
                                    {
                                        // Fallback to generic name if we can't get the process name
                                        processName = $"Session_{i}";
                                    }
                                }
                                catch (Exception ex) { 
                                    Log($"### Error getting session info for session {i}: {ex.Message}");
                                    processName = $"Session_{i}";
                                }
                                
                                // Check if this process is in our list of relevant processes
                                bool isRelevantProcess = processNameList.Any(targetProcess => 
                                    processName.Equals(targetProcess, StringComparison.OrdinalIgnoreCase));
                                
                                if (!isRelevantProcess)
                                {
                                    continue;
                                }
                                
                                // Only log when we find a relevant process (this is important info)
                                Log($"%%% Found relevant process: {processName} on device: {device.FriendlyName}");
                                
                                if (!newAudioSessionCache.ContainsKey(processName))
                                {
                                    newAudioSessionCache[processName] = [];
                                }
                                
                                var audioSessionInfo = new AudioSessionInfo
                                {
                                    ProcessName = processName,
                                    Session = session,
                                    DeviceName = device.FriendlyName,
                                    IsMuted = session.SimpleAudioVolume?.Mute ?? false
                                };
                                
                                newAudioSessionCache[processName].Add(audioSessionInfo);
                            }
                        }
                    }
                    catch (Exception) { /* Ignore individual device errors */ }
                    finally
                    {
                        device?.Dispose();
                    }
                }
                
                lock (audioCacheLock)
                {
                    audioSessionCache.Clear();
                    foreach (var kvp in newAudioSessionCache)
                    {
                        audioSessionCache[kvp.Key] = kvp.Value;
                    }
                }
                
                // Only log cache status if there are relevant processes found
                if (newAudioSessionCache.Count > 0)
                {
                    //Log($"%%% Audio cache updated: {newAudioSessionCache.Count} relevant processes found");
                }
            }
            catch (Exception ex)
            {
                Log($"### Error refreshing audio session cache: {ex.Message}");
            }
        }

        private void ProcessLogQueue()
        {
            while (logThreadRunning)
            {
                try
                {
                    string? logEntry = null;
                    lock (logLock)
                    {
                        if (logQueue.Count > 0)
                        {
                            logEntry = logQueue.Dequeue();
                        }
                    }

                    if (logEntry != null)
                    {
                        WriteLogToFile(logEntry);
                    }
                    else
                    {
                        Thread.Sleep(10); // Small delay when no logs to process
                    }
                }
                catch (Exception) { /* Ignore logging errors */ }
            }
        }

        private void WriteLogToFile(string logEntry)
        {
            string logPath = "GlobalHotkeys.log";
            try
            {
                File.AppendAllText(logPath, logEntry + Environment.NewLine);
                Console.WriteLine(logEntry);
            }
            catch (Exception) { /* Ignore file write errors */ }
        }

        ~InvisibleForm() {
            logThreadRunning = false;
            processCacheRefreshTimer?.Dispose();
            audioCacheRefreshTimer?.Dispose();
            audioDeviceEnumerator?.Dispose();
            hook.Dispose();
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
            processNameList = File.ReadAllText(applicationListPath).Split(new string[] { "\r\n", "\r", "\n" },StringSplitOptions.None).ToList();
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
            relevantSoundDevices = File.ReadAllText(soundDevicesPath).Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).ToList();
            foreach (string device in relevantSoundDevices)
            {
                Log($"$$$ Sound device scanned in: {device}");
            }
        }

        public void Log(params string[] logOutput)
        {
            string combinedLogOutput = $"{DateTime.Now:HH:mm:ss.fff}:\t{string.Join(" ", logOutput)}";
            
            // Queue log entry for background processing to avoid blocking
            lock (logLock)
            {
                logQueue.Enqueue(combinedLogOutput);
            }
            
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
                        // Pre-allocate list capacity for better performance
                        currentAppsToMute.Capacity = Math.Max(currentAppsToMute.Capacity, processNameList.Count);
                        
                        foreach (string processNameToLookFor in processNameList)
                        {
                            if (processCache.TryGetValue(processNameToLookFor, out Process[]? processesFound))
                            {
                                currentAppsToMute.Add(processNameToLookFor);
                                Log($"\t@@@ Game Found: '{processNameToLookFor}' (cached)");
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
                    if (processNameList.Contains(procNameTemp)) 
                    {
                        Log($"### Process: '{procNameTemp}' already exists in mute list, skipping");
                    }
                    else 
                    {
                        processNameList.Add(procNameTemp);
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
                        bool isRelevantDevice = relevantSoundDevices.Any(deviceName => 
                            device.FriendlyName.Contains(deviceName, StringComparison.OrdinalIgnoreCase));
                        
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
            Process proc = new();
            ProcessStartInfo pi = new()

            {
                UseShellExecute = true,
                FileName = @filePath,
                Arguments = string.Join(" ",arguments),
                CreateNoWindow = true
            };
            proc.StartInfo = pi;
            proc.Start();
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
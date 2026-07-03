using System;
using Microsoft.Win32;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Windows.Forms;

namespace WallpaperSlideshow
{
    class Program
    {
        static string sStop = "Stopping all instances.";
        static string sRequires = "This program requires Windows 8 or later.";
        static string sInstance = "Another instance is already running.";
        static string sFolder = "Folder does not exist: ";
        static string sWait = "Invalid wait time: ";
        static string sInvalidFolders = "One or more specified folders are invalid. Please check your settings.";
        static string sTitle = "WallpaperSlideshow";
        static string sHelp =
                        "Multi-monitor, multi-folder, wallpaper slideshow\n" +
                        "Automatically detects monitor changes and folder content changes\n" +
                        "Uses minimal memory and survives Explorer restart\n" +
                        "Supported image file formats: .jpg .jpeg .png .bmp\n" +
                        "Usage: wallpaperslideshow [folder1] [seconds1] [folder2] [seconds2] ...\n" +
                        "Where: folder1 = image folder for monitor 1\n" +
                        "Where: seconds1 = wait time between images for monitor 1 in seconds\n" +
                        "Settings are written to HKCU\\Software\\WallpaperSlideshow\n" +
                        "If no arguments are provided, settings from the registry will be used.\n" +
                        "Stop slideshow: wallpaperslideshow /x";

        static Mutex mutex = new Mutex(true, "{6B63C8F3-18F1-4D57-87F0-B6281CE747FA}");

        static bool Sync = false;
        static bool Shuffle = false;
        static bool ReShuffle = false;

        private static readonly Random rng = new Random();

        [DllImport("kernel32.dll")]
        static extern bool AttachConsole(int dwProcessId);

        [DllImport("kernel32.dll")]
        static extern bool FreeConsole();

        private const int ATTACH_PARENT_PROCESS = -1;

        [STAThread]
        static void Main(string[] args)
        {
            // Attach to parent console if launched from command line
            bool hasConsole = AttachConsole(ATTACH_PARENT_PROCESS);

            if (args.Length == 1)
            {
                string arg = args[0].ToLowerInvariant();
                if (arg == "/kill" || arg == "/stop" || arg == "/exit" || arg == "/quit" || arg == "/x")
                {
                    ConWrite(sStop, hasConsole);
                    KillAllInstances();
                    return;
                }
                if (arg == "/help" || arg == "/?")
                {
                    ConWrite(sHelp, hasConsole);
                    return;
                }
            }

            // Windows 8+
            string NTVer = (string)Registry.GetValue(
                @"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion",
                "CurrentVersion", "6.0");

            if (Version.Parse(NTVer) < new Version("6.2"))
            {
                ConWrite(sRequires, hasConsole);
                return;
            }

            // Ensure only one instance runs at a time
            if (!mutex.WaitOne(TimeSpan.Zero, true))
            {
                ConWrite(sInstance, hasConsole);
                return;
            }

            using (RegistryKey baseKey = Registry.CurrentUser.OpenSubKey(@"Software\WallpaperSlideshow"))
            {
                if (baseKey != null)
                {
                    Sync      = (int)(baseKey.GetValue("Sync",      0) ?? 0) != 0;
                    Shuffle   = (int)(baseKey.GetValue("Shuffle",   0) ?? 0) != 0;
                    ReShuffle = (int)(baseKey.GetValue("ReShuffle", 0) ?? 0) != 0;
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Load folder/wait pairs from registry
            List<FolderWait> folderWaits = LoadFromRegistry();

            // Override/add with command-line arguments if present
            if (args.Length >= 2 && args.Length % 2 == 0)
            {
                for (int i = 0; i < args.Length; i += 2)
                {
                    string folder = args[i];
                    if (!Directory.Exists(folder))
                    {
                        ConWrite($"{sFolder}{folder}", hasConsole);
                        return;
                    }

                    if (!int.TryParse(args[i + 1], out int wait) || wait <= 0)
                    {
                        ConWrite($"{sWait}{args[i + 1]}", hasConsole);
                        return;
                    }

                    if (i / 2 < folderWaits.Count)
                        folderWaits[i / 2] = new FolderWait(folder, wait); // override
                    else
                        folderWaits.Add(new FolderWait(folder, wait));       // add
                }

                // Save updated folder/wait pairs to registry
                using (var baseKey = Registry.CurrentUser.CreateSubKey(@"Software\WallpaperSlideshow"))
                {
                    for (int i = 0; i < folderWaits.Count; i++)
                    {
                        var fw = folderWaits[i];
                        using (var subKey = baseKey.CreateSubKey(i.ToString()))
                        {
                            subKey.SetValue("", fw.Folder, RegistryValueKind.String);
                            subKey.SetValue("Wait", fw.Wait, RegistryValueKind.DWord);
                        }
                    }
                }
            }

            // Detach from console for normal operation (runs hidden)
            if (hasConsole)
            {
                FreeConsole();
            }

            // Initialize COM wallpaper handler
            IDesktopWallpaper handler = (IDesktopWallpaper)new DesktopWallpaperClass();
            handler.SetPosition(4); // Fill

            // Start slideshow loop
            RunSlideshowLoop(handler, folderWaits);
        }

        static void ConWrite(string text, bool hasConsole = false)
        {
            if (hasConsole)
            {
                Console.WriteLine($"\n\n{text}");
                SendKeys.SendWait("{ENTER}");
            }
        }


        static void RunSlideshowLoop(IDesktopWallpaper handler, List<FolderWait> initialFolderWaits)
        {
            Random random = new Random();
            List<FolderWait> folderWaits = initialFolderWaits ?? LoadFromRegistry();
            if (!AreFoldersValid(folderWaits))
            {
                MessageBox.Show(sInvalidFolders, sTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Dictionary<string, int> monitorIndexes = new Dictionary<string, int>();
            Dictionary<string, string[]> monitorImages = new Dictionary<string, string[]>();
            Dictionary<string, int> monitorWaits = new Dictionary<string, int>();
            Dictionary<string, DateTime> nextChangeTime = new Dictionary<string, DateTime>();
            Dictionary<string, int> monitorFolderIndexes = new Dictionary<string, int>(); // Track which folder enumeration index each monitor uses

            // FileSystemWatcher tracking
            Dictionary<string, FileSystemWatcher> folderWatchers = new Dictionary<string, FileSystemWatcher>();
            HashSet<string> foldersWithChanges = new HashSet<string>();
            object lockObj = new object();

            List<FolderWait> lastFolderWaits = new List<FolderWait>(folderWaits);
            DateTime lastGcTime = DateTime.Now;
            DateTime lastLoopTime = DateTime.Now;

            while (true)
            {
                try
                {
                    Thread.Sleep(1000); // check every second

                    // Periodic GC every 5 minutes
                    if ((DateTime.Now - lastGcTime).TotalMinutes >= 5)
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        lastGcTime = DateTime.Now;
                    }

                    bool folderWaitsChanged = false;

                    // Wake-from-sleep detection: if more than 3 seconds elapsed, system likely resumed from sleep
                    DateTime currentLoopTime = DateTime.Now;
                    bool wakeFromSleep = (currentLoopTime - lastLoopTime).TotalSeconds > 3;
                    lastLoopTime = currentLoopTime;

                    if (wakeFromSleep)
                    {
                        // Wait 2 seconds for all monitors to come back on before doing anything
                        Thread.Sleep(2000);
                    }

                    List<FolderWait> newFolderWaits = LoadFromRegistry();
                    if (!FolderWaitsEqual(lastFolderWaits, newFolderWaits) && AreFoldersValid(newFolderWaits))
                    {
                        folderWaits = newFolderWaits;
                        lastFolderWaits = new List<FolderWait>(newFolderWaits);
                        folderWaitsChanged = true;
                    }

                    uint monitorCount = handler.GetMonitorDevicePathCount();
                    HashSet<string> currentMonitors = new HashSet<string>();
                    HashSet<string> activeFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    // Pre-pass: identify folder paths used by more than one monitor
                    HashSet<string> sharedFolderPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    {
                        Dictionary<string, int> pathUsageCount = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                        for (uint j = 0; j < monitorCount; j++)
                        {
                            string tempId;
                            handler.GetMonitorDevicePathAt(j, out tempId);
                            if (string.IsNullOrEmpty(tempId)) continue;
                            int tempIdx = (int)j < folderWaits.Count ? (int)j : 0;
                            string tempPath = folderWaits[tempIdx].Folder;
                            pathUsageCount[tempPath] = pathUsageCount.ContainsKey(tempPath) ? pathUsageCount[tempPath] + 1 : 1;
                        }
                        foreach (var kv in pathUsageCount)
                            if (kv.Value > 1) sharedFolderPaths.Add(kv.Key);
                    }

                    // Track start indexes already assigned per shared folder path to ensure exclusive random selection
                    Dictionary<string, List<int>> folderAssignedIndexes = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

                    // Folders whose change notifications will be cleared after all monitors have been processed
                    HashSet<string> processedFolderChanges = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    for (uint i = 0; i < monitorCount; i++)
                    {
                        string monitorId;
                        handler.GetMonitorDevicePathAt(i, out monitorId);

                        // Skip monitors with invalid/empty MonitorIDs (phantom monitors)
                        if (string.IsNullOrEmpty(monitorId))
                        {
                            continue;
                        }

                        currentMonitors.Add(monitorId);

                        int idx = (int)i < folderWaits.Count ? (int)i : 0;
                        string folder = folderWaits[idx].Folder;
                        int wait = folderWaits[idx].Wait;

                        monitorFolderIndexes[monitorId] = idx; // Track which folder enumeration index each monitor uses
                        activeFolders.Add(folder); // Track folders currently assigned to monitors
                        monitorWaits[monitorId] = wait;

                        bool needsReload = false;

                        // Check if folder contents changed
                        lock (lockObj)
                        {
                            if (foldersWithChanges.Contains(folder))
                            {
                                processedFolderChanges.Add(folder);
                                needsReload = true;
                            }
                        }

                        // Load images only if monitor is new, folder/wait changed, or folder contents changed
                        if (!monitorImages.ContainsKey(monitorId) || folderWaitsChanged || needsReload)
                        {
                            if (!Directory.Exists(folder)) continue;

                            monitorImages[monitorId] = Directory.GetFiles(folder).Where(IsImage).ToArray();
                            string[] images = monitorImages[monitorId];

                            if (images.Length > 0)
                            {
                                if (Shuffle) ShuffleArray(images);

                                int startIdx;

                                if (Shuffle)
                                {
                                    // Array was just shuffled, so any saved index refers to
                                    // a different image - always start at the beginning
                                    startIdx = 0;
                                }
                                else if (Sync)
                                {
                                    // For sync mode, all monitors start at the same image
                                    startIdx = 0;
                                }
                                else if (sharedFolderPaths.Contains(folder))
                                {
                                    // Shared folder path, Sync=off: pick an exclusive random index so
                                    // monitors sharing a folder don't accidentally start in sync
                                    if (!folderAssignedIndexes.ContainsKey(folder))
                                        folderAssignedIndexes[folder] = new List<int>();

                                    List<int> assigned = folderAssignedIndexes[folder];
                                    if (assigned.Count < images.Length)
                                    {
                                        // Pick randomly from indexes not yet assigned to another monitor
                                        List<int> available = new List<int>();
                                        for (int k = 0; k < images.Length; k++)
                                            if (!assigned.Contains(k)) available.Add(k);
                                        startIdx = available[random.Next(available.Count)];
                                    }
                                    else
                                    {
                                        // More monitors than images; all indexes taken, just pick random
                                        startIdx = random.Next(images.Length);
                                    }
                                    assigned.Add(startIdx);
                                }
                                else
                                {
                                    // Try to load the saved index from registry
                                    int savedIdx = LoadImageIndexFromRegistry(idx, -1);

                                    if (savedIdx >= 0 && savedIdx < images.Length)
                                    {
                                        // Use saved index if valid
                                        startIdx = savedIdx;
                                    }
                                    else
                                    {
                                        // Otherwise, set random starting image
                                        startIdx = random.Next(images.Length);
                                    }
                                }

                                // Immediately apply the saved wallpaper, then queue the next image
                                if (File.Exists(images[startIdx]))
                                {
                                    try { handler.SetWallpaper(monitorId, images[startIdx]); }
                                    catch { }
                                }
                                monitorIndexes[monitorId] = (startIdx + 1) % images.Length;
                                SaveImageIndexToRegistry(idx, startIdx);
                            }

                            // Restore saved next change time, clamped so it is never more than one full wait period from now
                            DateTime savedNextChange = LoadNextChangeTimeFromRegistry(idx, DateTime.Now.AddSeconds(wait));
                            DateTime maxNextChange = DateTime.Now.AddSeconds(wait);
                            nextChangeTime[monitorId] = savedNextChange < maxNextChange ? savedNextChange : maxNextChange;
                        }

                        // Initialize index and nextChangeTime for new monitors
                        if (!monitorIndexes.ContainsKey(monitorId))
                            monitorIndexes[monitorId] = 0;
                        if (!nextChangeTime.ContainsKey(monitorId))
                            nextChangeTime[monitorId] = DateTime.Now;
                    }

                    // Now that all monitors have seen the folder changes, clear them
                    if (processedFolderChanges.Count > 0)
                    {
                        lock (lockObj)
                        {
                            foreach (string f in processedFolderChanges)
                                foldersWithChanges.Remove(f);
                        }
                    }

                    // Update FileSystemWatchers: remove watchers for folders no longer in use
                    foreach (string watchedFolder in folderWatchers.Keys.ToList())
                    {
                        if (!activeFolders.Contains(watchedFolder))
                        {
                            folderWatchers[watchedFolder].Dispose();
                            folderWatchers.Remove(watchedFolder);
                        }
                    }

                    // Add watchers for new folders
                    foreach (string folder in activeFolders)
                    {
                        if (!folderWatchers.ContainsKey(folder) && Directory.Exists(folder))
                        {
                            try
                            {
                                FileSystemWatcher watcher = new FileSystemWatcher(folder);
                                watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
                                watcher.Filter = "*.*";

                                FileSystemEventHandler handler_event = (sender, e) =>
                                {
                                    if (IsImage(e.FullPath))
                                    {
                                        lock (lockObj)
                                        {
                                            foldersWithChanges.Add(folder); // Mark folder for reload on next loop iteration
                                        }
                                    }
                                };

                                RenamedEventHandler renamed_event = (sender, e) =>
                                {
                                    if (IsImage(e.FullPath) || IsImage(e.OldFullPath))
                                    {
                                        lock (lockObj)
                                        {
                                            foldersWithChanges.Add(folder);
                                        }
                                    }
                                };

                                watcher.Created += handler_event;
                                watcher.Deleted += handler_event;
                                watcher.Renamed += renamed_event;
                                watcher.Changed += handler_event;

                                watcher.EnableRaisingEvents = true;
                                folderWatchers[folder] = watcher;
                            }
                            catch { } // Ignore watcher creation failures
                        }
                    }

                    DateTime now = DateTime.Now;

                    // On wake from sleep, immediately re-apply the current wallpaper to all monitors
                    if (wakeFromSleep)
                    {
                        foreach (var monitorId in currentMonitors)
                        {
                            if (!monitorImages.ContainsKey(monitorId) || !monitorIndexes.ContainsKey(monitorId))
                                continue;

                            string[] images = monitorImages[monitorId];
                            // monitorIndexes holds the *next* image to show, so step back one to get the current
                            int currentIdx = (monitorIndexes[monitorId] - 1 + images.Length) % images.Length;
                            if (images.Length > 0 && File.Exists(images[currentIdx]))
                            {
                                try { handler.SetWallpaper(monitorId, images[currentIdx]); }
                                catch { }
                            }
                        }
                    }

                    foreach (var monitorId in currentMonitors)
                    {
                        if (!nextChangeTime.ContainsKey(monitorId) || !monitorImages.ContainsKey(monitorId))
                            continue;

                        if (now >= nextChangeTime[monitorId])
                        {
                            string[] images = monitorImages[monitorId];
                            if (images.Length > 0)
                            {
                                int idxNext = monitorIndexes[monitorId];
                                if (File.Exists(images[idxNext]))
                                {
                                    try 
                                    { 
                                        handler.SetWallpaper(monitorId, images[idxNext]);
                                        if (monitorFolderIndexes.ContainsKey(monitorId))
                                        {
                                            SaveImageIndexToRegistry(monitorFolderIndexes[monitorId], idxNext);
                                        }
                                    }
                                    catch { }
                                }
                                int nextIdx = (idxNext + 1) % images.Length;

                                // If we've just shown the last image and ReShuffle is on, reshuffle before wrapping
                                if (ReShuffle && nextIdx == 0)
                                    ShuffleArray(images);

                                monitorIndexes[monitorId] = nextIdx;
                                nextChangeTime[monitorId] = now.AddSeconds(monitorWaits[monitorId]);
                                if (monitorFolderIndexes.ContainsKey(monitorId))
                                {
                                    SaveNextChangeTimeToRegistry(monitorFolderIndexes[monitorId], nextChangeTime[monitorId]);
                                }
                            }
                        }
                    }

                    // Remove monitors that disappeared
                    foreach (string key in new List<string>(monitorIndexes.Keys))
                    {
                        if (!currentMonitors.Contains(key))
                        {
                            monitorIndexes.Remove(key);
                            monitorImages.Remove(key);
                            monitorWaits.Remove(key);
                            nextChangeTime.Remove(key);
                            monitorFolderIndexes.Remove(key);
                        }
                    }
                }
                catch (COMException)
                {
                    // COM object became invalid (e.g., Explorer restart) - reinitialize
                    Thread.Sleep(2000); // Wait for Explorer to stabilize
                    try
                    {
                        handler = (IDesktopWallpaper)new DesktopWallpaperClass();
                        handler.SetPosition(4); // Fill
                    }
                    catch
                    {
                        Thread.Sleep(5000); // If reinitialization fails, wait longer
                    }
                }
                catch
                {
                    // Other unexpected errors - wait and continue
                    Thread.Sleep(1000);
                }
            }
        }

        private static void ShuffleArray<T>(T[] array)
        {
            for (int i = array.Length - 1; i > 0; i--)
            {
                int j = rng.Next(i + 1);
                (array[i], array[j]) = (array[j], array[i]);
            }
        }

        static bool IsImage(string f)
        {
            string ext = Path.GetExtension(f).ToLowerInvariant();
            return ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".bmp";
        }

        static void KillAllInstances()
        {
            int currentId = Process.GetCurrentProcess().Id;
            Process[] processes = Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName);

            foreach (Process p in processes)
            {
                if (p.Id != currentId)
                {
                    try { p.Kill(); } catch { }
                }
                p.Dispose();
            }
        }

        static bool AreFoldersValid(List<FolderWait> list)
        {
            if (list.Count == 0) return false;
            foreach (FolderWait fw in list)
            {
                if (string.IsNullOrWhiteSpace(fw.Folder) || !Directory.Exists(fw.Folder))
                    return false;
            }
            return true;
        }

        static bool FolderWaitsEqual(List<FolderWait> a, List<FolderWait> b)
        {
            if (a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++)
                if (a[i].Folder != b[i].Folder || a[i].Wait != b[i].Wait)
                    return false;
            return true;
        }

        static List<FolderWait> LoadFromRegistry()
        {
            List<FolderWait> result = new List<FolderWait>();
            using (RegistryKey baseKey = Registry.CurrentUser.OpenSubKey(@"Software\WallpaperSlideshow"))
            {
                if (baseKey == null) return result;

                int idx = 0;
                while (true)
                {
                    using (RegistryKey subKey = baseKey.OpenSubKey(idx.ToString()))
                    {
                        if (subKey == null) break;

                        string folder = subKey.GetValue("") as string;
                        int wait = (int)(subKey.GetValue("Wait") ?? 60);
                        result.Add(new FolderWait(folder, wait));
                    }
                    idx++;
                }
            }
            return result;
        }

        static void SaveImageIndexToRegistry(int folderIndex, int imageIndex)
        {
            try
            {
                using (RegistryKey baseKey = Registry.CurrentUser.CreateSubKey(@"Software\WallpaperSlideshow"))
                {
                    using (RegistryKey subKey = baseKey.CreateSubKey(folderIndex.ToString()))
                    {
                        subKey.SetValue("CurrentImageIndex", imageIndex, RegistryValueKind.DWord);
                    }
                }
            }
            catch { }
        }

        static int LoadImageIndexFromRegistry(int folderIndex, int defaultIndex)
        {
            try
            {
                using (RegistryKey baseKey = Registry.CurrentUser.OpenSubKey(@"Software\WallpaperSlideshow"))
                {
                    if (baseKey == null) return defaultIndex;
                    using (RegistryKey subKey = baseKey.OpenSubKey(folderIndex.ToString()))
                    {
                        if (subKey == null) return defaultIndex;
                        object value = subKey.GetValue("CurrentImageIndex");
                        if (value != null)
                        {
                            return (int)value;
                        }
                    }
                }
            }
            catch { }
            return defaultIndex;
        }

        static void SaveNextChangeTimeToRegistry(int folderIndex, DateTime nextChangeTime)
        {
            try
            {
                using (RegistryKey baseKey = Registry.CurrentUser.CreateSubKey(@"Software\WallpaperSlideshow"))
                {
                    using (RegistryKey subKey = baseKey.CreateSubKey(folderIndex.ToString()))
                    {
                        subKey.SetValue("NextChangeTime", nextChangeTime.ToBinary(), RegistryValueKind.QWord);
                    }
                }
            }
            catch { }
        }

        static DateTime LoadNextChangeTimeFromRegistry(int folderIndex, DateTime defaultTime)
        {
            try
            {
                using (RegistryKey baseKey = Registry.CurrentUser.OpenSubKey(@"Software\WallpaperSlideshow"))
                {
                    if (baseKey == null) return defaultTime;
                    using (RegistryKey subKey = baseKey.OpenSubKey(folderIndex.ToString()))
                    {
                        if (subKey == null) return defaultTime;
                        object value = subKey.GetValue("NextChangeTime");
                        if (value != null)
                        {
                            return DateTime.FromBinary((long)value);
                        }
                    }
                }
            }
            catch { }
            return defaultTime;
        }
    }

    class FolderWait
    {
        public string Folder { get; set; }
        public int Wait { get; set; }
        public FolderWait(string folder, int wait) { Folder = folder; Wait = wait; }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct Rect { public int Left, Top, Right, Bottom; }

    [ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDesktopWallpaper
    {
        void SetWallpaper(string monitorID, string imagePath);
        string GetWallpaper(string monitorID);
        void GetMonitorDevicePathAt(uint monitorIndex, [Out, MarshalAs(UnmanagedType.LPWStr)] out string monitorID);
        uint GetMonitorDevicePathCount();
        Rect GetMonitorRECT(string monitorID);
        void SetBackgroundColor(uint color);
        uint GetBackgroundColor();
        void SetPosition(int position);
        string GetPosition();
        void SetSlideshow(IntPtr items);
        IntPtr GetSlideshow();
        bool Enable();
    }

    [ComImport, Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD")]
    public class DesktopWallpaperClass { }
}
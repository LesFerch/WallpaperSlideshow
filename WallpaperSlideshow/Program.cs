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
                        "If no arguments are provided, settings from the registry will be used.";

        static Mutex mutex = new Mutex(true, "{6B63C8F3-18F1-4D57-87F0-B6281CE747FA}");

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
                    ConWrite(sStop);
                    KillAllInstances();
                    return;
                }
                if (arg == "/help" || arg == "/?")
                {
                    ConWrite(sHelp);
                    return;
                }
            }

            // Windows 8+
            string NTVer = (string)Registry.GetValue(
                @"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion",
                "CurrentVersion", "6.0");

            if (Version.Parse(NTVer) < new Version("6.2"))
            {
                ConWrite(sRequires);
                return;
            }

            // Ensure only one instance runs at a time
            if (!mutex.WaitOne(TimeSpan.Zero, true))
            {
                ConWrite(sInstance);
                return; 
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
                        ConWrite($"{sFolder}{folder}");
                        return; 
                    }

                    if (!int.TryParse(args[i + 1], out int wait) || wait <= 0)
                    {
                        ConWrite($"{sWait}{args[i + 1]}"); 
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

        static void ConWrite(string text)
        {
            Console.WriteLine($"\n\n{text}");
            SendKeys.SendWait("{ENTER}");
        }   

        static void RunSlideshowLoop(IDesktopWallpaper handler, List<FolderWait> initialFolderWaits)
        {
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

            // FileSystemWatcher tracking
            Dictionary<string, FileSystemWatcher> folderWatchers = new Dictionary<string, FileSystemWatcher>();
            HashSet<string> foldersWithChanges = new HashSet<string>();
            object lockObj = new object();

            List<FolderWait> lastFolderWaits = new List<FolderWait>(folderWaits);
            DateTime lastGcTime = DateTime.Now;

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

                    for (uint i = 0; i < monitorCount; i++)
                    {
                        string monitorId;
                        handler.GetMonitorDevicePathAt(i, out monitorId);
                        currentMonitors.Add(monitorId);

                        int idx = (int)i < folderWaits.Count ? (int)i : 0;
                        string folder = folderWaits[idx].Folder;
                        int wait = folderWaits[idx].Wait;

                        activeFolders.Add(folder); // Track folders currently assigned to monitors
                        monitorWaits[monitorId] = wait;

                        bool needsReload = false;

                        // Check if folder contents changed
                        lock (lockObj)
                        {
                            if (foldersWithChanges.Contains(folder))
                            {
                                foldersWithChanges.Remove(folder);
                                needsReload = true;
                            }
                        }

                        // Load images only if monitor is new, folder/wait changed, or folder contents changed
                        if (!monitorImages.ContainsKey(monitorId) || folderWaitsChanged || needsReload)
                        {
                            if (!Directory.Exists(folder)) continue;
                            monitorImages[monitorId] = Directory.GetFiles(folder).Where(IsImage).ToArray();
                            monitorIndexes[monitorId] = 0;
                            nextChangeTime[monitorId] = DateTime.Now;
                        }

                        // Initialize index and nextChangeTime for new monitors
                        if (!monitorIndexes.ContainsKey(monitorId))
                            monitorIndexes[monitorId] = 0;
                        if (!nextChangeTime.ContainsKey(monitorId))
                            nextChangeTime[monitorId] = DateTime.Now;
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
                                    try { handler.SetWallpaper(monitorId, images[idxNext]); }
                                    catch { }
                                }
                                monitorIndexes[monitorId] = (idxNext + 1) % images.Length;
                                nextChangeTime[monitorId] = now.AddSeconds(monitorWaits[monitorId]);
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
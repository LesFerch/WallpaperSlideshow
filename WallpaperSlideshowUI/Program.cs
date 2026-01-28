using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.IO;
using System.Text;

namespace WallpaperSlideshowUI
{
    class Program
    {
        static float ScaleFactor = GetScale();
        static bool Dark = isDark();
        static string startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        static string shortcutPath = System.IO.Path.Combine(startupFolder, "WallpaperSlideshow.lnk");
        static string myPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
        static bool IsStartup = IsStartupShortcutExists();
        static string sStartup = "Run at startup";
        static string sSetWait = "Set Wait Time";
        static string sHelp = "Help";
        static string sExit = "Exit";
        static string sOK = "OK";
        static string sHours = "Hours";
        static string sMinutes = "Minutes";
        static string sSeconds = "Seconds";
        static string sMonitor = "Monitor";
        static string sWait = "Wait";
        static string sFolder = "Slideshow Folder";

        // Prevents multiple instances of the settings UI from running simultaneously
        static Mutex mutex = new Mutex(true, "{E41A231A-A858-41DF-8F0E-A5A29F40301C}");

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        private const int SW_RESTORE = 9;

        [STAThread]
        static void Main(string[] args)
        {
            // Do not allow multiple instances
            if (!mutex.WaitOne(TimeSpan.Zero, true))
            {
                BringExistingInstanceToFront();
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            LoadLanguageStrings();

            var folderWaits = LoadFromRegistry();
            using (var dlg = new SlideshowSettingsForm(folderWaits))
            {
                dlg.ShowDialog();
            }
        }

        private static void BringExistingInstanceToFront()
        {
            var currentProcess = Process.GetCurrentProcess();
            var processes = Process.GetProcessesByName(currentProcess.ProcessName);

            foreach (var process in processes)
            {
                if (process.Id != currentProcess.Id && process.MainWindowHandle != IntPtr.Zero)
                {
                    IntPtr handle = process.MainWindowHandle;

                    // If minimized, restore it
                    if (IsIconic(handle))
                    {
                        ShowWindow(handle, SW_RESTORE);
                    }

                    // Bring to foreground
                    SetForegroundWindow(handle);
                    break;
                }
            }

            foreach (var process in processes)
            {
                process.Dispose();
            }
        }

        // Load language strings from INI file
        static void LoadLanguageStrings()
        {
            string iniFile = System.IO.Path.Combine(myPath, "language.ini");

            if (!File.Exists(iniFile)) return;

            string lang = GetLang();

            sMonitor = ReadString(iniFile, lang, "sMonitor", sMonitor);
            sWait = ReadString(iniFile, lang, "sWait", sWait);
            sFolder = ReadString(iniFile, lang, "sFolder", sFolder);
            sOK = ReadString(iniFile, lang, "sOK", sOK);
            sHelp = ReadString(iniFile, lang, "sHelp", sHelp);
            sExit = ReadString(iniFile, lang, "sExit", sExit);
            sStartup = ReadString(iniFile, lang, "sStartup", sStartup);
            sSetWait = ReadString(iniFile, lang, "sSetWait", sSetWait);
            sHours = ReadString(iniFile, lang, "sHours", sHours);
            sMinutes = ReadString(iniFile, lang, "sMinutes", sMinutes);
            sSeconds = ReadString(iniFile, lang, "sSeconds", sSeconds);
        }


        static string ReadString(string iniFile, string section, string key, string defaultValue)
        {
            try
            {
                if (File.Exists(iniFile))
                {
                    return IniFileParser.ReadValue(section, key, defaultValue, iniFile);
                }
            }
            catch { }

            return defaultValue;
        }

        // INI file parser
        public static class IniFileParser
        {
            public static string ReadValue(string section, string key, string defaultValue, string filePath)
            {
                try
                {
                    var lines = File.ReadAllLines(filePath, Encoding.UTF8);
                    string currentSection = null;

                    foreach (var line in lines)
                    {
                        string trimmedLine = line.Trim();

                        if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
                        {
                            currentSection = trimmedLine.Substring(1, trimmedLine.Length - 2);
                        }
                        else if (currentSection == section)
                        {
                            var parts = trimmedLine.Split(new char[] { '=' }, 2);
                            if (parts.Length == 2 && parts[0].Trim() == key)
                            {
                                return parts[1].Trim();
                            }
                        }
                    }
                }
                catch (Exception)
                {
                }
                return defaultValue;
            }
        }


        // Get language from INI file or system
        static string GetLang()
        {
            string iniFile = System.IO.Path.Combine(myPath, "WallpaperSlideshow.ini");

            string lang = ReadString(iniFile, "General", "Lang", "");
            if (lang != "") return lang;

            lang = "en";

            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\International");
                if (key != null)
                {
                    lang = key.GetValue("LocaleName") as string;
                    key.Close();
                }
            }
            catch { }

            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop");
                if (key != null)
                {
                    string[] preferredLanguages = key.GetValue("PreferredUILanguages") as string[];
                    if (preferredLanguages != null && preferredLanguages.Length > 0)
                    {
                        lang = preferredLanguages[0];
                    }
                    key.Close();
                }
            }
            catch { }

            return lang.Substring(0, 2).ToLower();
        }

        // Loads settings from registry: HKCU\Software\WallpaperSlideshow\0, \1, \2, etc.
        // Each subkey contains: (Default) = folder path, Wait = seconds (DWORD)
        public static List<(string folder, int wait)> LoadFromRegistry()
        {
            var result = new List<(string folder, int wait)>();
            using (var baseKey = Registry.CurrentUser.OpenSubKey(@"Software\WallpaperSlideshow"))
            {
                if (baseKey == null) return result;
                int idx = 0;
                while (true)
                {
                    using (var subKey = baseKey.OpenSubKey(idx.ToString()))
                    {
                        if (subKey == null) break;
                        string folder = subKey.GetValue("") as string;
                        int wait = (int)(subKey.GetValue("Wait") ?? 60);
                        result.Add((folder, wait));
                    }
                    idx++;
                }
            }
            return result;
        }

        static float GetScale()
        {
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiX = graphics.DpiX;
                return dpiX / 96;
            }
        }

        // Checks Windows registry to determine if dark mode is enabled
        public static bool isDark()
        {
            const string keyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            const string valueName = "AppsUseLightTheme";

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath))
            {
                if (key != null)
                {
                    object value = key.GetValue(valueName);
                    if (value is int intValue)
                    {
                        return intValue == 0;
                    }
                }
            }
            return false;
        }

        // Check if startup shortcut exists
        public static bool IsStartupShortcutExists()
        {
            return System.IO.File.Exists(shortcutPath);
        }

        // Create startup shortcut
        public static void CreateStartupShortcut()
        {
            string exePath = System.IO.Path.Combine(myPath,"WallpaperSlideshow.exe");

            if (!System.IO.File.Exists(exePath)) return;

            // Create shortcut using Windows Script Host
            Type shellType = Type.GetTypeFromProgID("WScript.Shell");
            dynamic shell = Activator.CreateInstance(shellType);
            var shortcut = shell.CreateShortcut(shortcutPath);
            shortcut.TargetPath = exePath;
            shortcut.WorkingDirectory = System.IO.Path.GetDirectoryName(exePath);
            shortcut.Description = "WallpaperSlideshow";
            shortcut.Save();
            Marshal.ReleaseComObject(shortcut);
            Marshal.ReleaseComObject(shell);
        }

        // Remove startup shortcut
        public static void RemoveStartupShortcut()
        {
            if (System.IO.File.Exists(shortcutPath))
            {
                System.IO.File.Delete(shortcutPath);
            }
        }

        public enum DWMWINDOWATTRIBUTE : uint
        {
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20,
        }

        [DllImport("dwmapi.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        public static extern void DwmSetWindowAttribute(IntPtr hwnd, DWMWINDOWATTRIBUTE attribute, ref int pvAttribute, uint cbAttribute);

        public static void DarkTitleBar(IntPtr hWnd)
        {
            var preference = Convert.ToInt32(true);
            DwmSetWindowAttribute(hWnd, DWMWINDOWATTRIBUTE.DWMWA_USE_IMMERSIVE_DARK_MODE, ref preference, sizeof(uint));
        }


        // Time picker dialog with hours, minutes, and seconds input
        public class TimePickerDialog : Form
        {
            private NumericUpDown numHours;
            private NumericUpDown numMinutes;
            private NumericUpDown numSeconds;
            private Button btnOK;
            private Label lblHours;
            private Label lblMinutes;
            private Label lblSeconds;

            public int TotalSeconds { get; private set; }

            public TimePickerDialog(int initialSeconds)
            {
                float scale = ScaleFactor;
                bool isDark = Dark;

                this.Text = sSetWait;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.StartPosition = FormStartPosition.CenterParent;
                this.ClientSize = new Size((int)(280 * scale), (int)(120 * scale));
                this.Font = new Font(this.Font.FontFamily, this.Font.Size + 2);
                this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

                // Convert total seconds to hours, minutes, seconds for display
                int hours = initialSeconds / 3600;
                int minutes = (initialSeconds % 3600) / 60;
                int seconds = initialSeconds % 60;

                lblHours = new Label { Text = sHours, Left = (int)(20 * scale), Top = (int)(20 * scale), Width = (int)(60 * scale) };
                numHours = new NumericUpDown { Left = (int)(90 * scale), Top = (int)(18 * scale), Width = (int)(60 * scale), Minimum = 0, Maximum = 99, Value = hours };

                lblMinutes = new Label { Text = sMinutes, Left = (int)(20 * scale), Top = (int)(50 * scale), Width = (int)(60 * scale) };
                numMinutes = new NumericUpDown { Left = (int)(90 * scale), Top = (int)(48 * scale), Width = (int)(60 * scale), Minimum = 0, Maximum = 59, Value = minutes };

                lblSeconds = new Label { Text = sSeconds, Left = (int)(20 * scale), Top = (int)(80 * scale), Width = (int)(60 * scale) };
                numSeconds = new NumericUpDown { Left = (int)(90 * scale), Top = (int)(78 * scale), Width = (int)(60 * scale), Minimum = 0, Maximum = 59, Value = seconds };

                btnOK = new Button
                {
                    Text = sOK,
                    DialogResult = DialogResult.OK,
                    Left = (int)(170 * scale),
                    Top = (int)(38 * scale),
                    Width = (int)(90 * scale),
                    Height = (int)(30 * scale)
                };

                if (isDark)
                {
                    this.BackColor = Color.FromArgb(32, 32, 32);
                    foreach (var lbl in new[] { lblHours, lblMinutes, lblSeconds })
                    {
                        lbl.ForeColor = Color.White;
                    }
                    foreach (var num in new[] { numHours, numMinutes, numSeconds })
                    {
                        num.BackColor = Color.FromArgb(60, 60, 60);
                        num.ForeColor = Color.White;
                    }
                    btnOK.FlatStyle = FlatStyle.Flat;
                    btnOK.FlatAppearance.BorderColor = SystemColors.Highlight;
                    btnOK.FlatAppearance.BorderSize = 1;
                    btnOK.BackColor = Color.FromArgb(60, 60, 60);
                    btnOK.FlatAppearance.MouseOverBackColor = Color.Black;
                    btnOK.ForeColor = Color.White;
                    DarkTitleBar(this.Handle);
                }

                btnOK.Click += (s, e) =>
                {
                    // Convert back to total seconds
                    TotalSeconds = (int)(numHours.Value * 3600 + numMinutes.Value * 60 + numSeconds.Value);
                    if (TotalSeconds == 0) TotalSeconds = 1; // Minimum 1 second
                };

                this.Controls.AddRange(new Control[] { lblHours, numHours, lblMinutes, numMinutes, lblSeconds, numSeconds, btnOK });
                this.AcceptButton = btnOK;
            }
        }

        // Main settings form with one row per monitor
        public class SlideshowSettingsForm : Form
        {
            private const int MAX_WAIT_SECONDS = 359999; // 99 hours, 59 minutes, 59 seconds
            private DataGridView grid;
            private Button btnOK;
            private Button btnHelp;
            private Button btnExit;
            private ToggleSwitch toggleStartup;
            private Label lblStartup;
            private List<(string folder, int wait)> folderWaits;
            private int monitorCount;
            private int hoveredButtonRow = -1;
            private int hoveredButtonCol = -1;

            public SlideshowSettingsForm(List<(string folder, int wait)> folderWaits)
            {
                this.folderWaits = folderWaits;

                // Get monitor count from System.Windows.Forms.Screen
                this.monitorCount = Screen.AllScreens.Length;

                float scale = ScaleFactor;

                this.Text = "WallpaperSlideshow";
                this.Width = (int)(800 * scale);
                this.Font = new Font(this.Font.FontFamily, this.Font.Size + 2);
                this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

                grid = new DataGridView
                {
                    AllowUserToAddRows = false,
                    RowHeadersVisible = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                    AllowUserToResizeRows = false,
                    AllowUserToResizeColumns = false,
                    AllowUserToOrderColumns = false,
                    ScrollBars = ScrollBars.None,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                    ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                    BorderStyle = BorderStyle.None
                };

                // Column layout: Monitor | ⏱ | Wait | 📁 | Slideshow Folder
                var monitorCol = new DataGridViewTextBoxColumn
                {
                    Name = "Monitor",
                    HeaderText = sMonitor,
                    Width = (int)(70 * scale),
                    ReadOnly = true
                };
                var timePickerCol = new DataGridViewButtonColumn
                {
                    Name = "TimePicker",
                    HeaderText = "",
                    Text = "⏱",
                    UseColumnTextForButtonValue = true,
                    Width = (int)(40 * scale)
                };
                var waitCol = new DataGridViewTextBoxColumn
                {
                    Name = "Wait",
                    HeaderText = sWait,
                    Width = (int)(80 * scale),
                    ReadOnly = false
                };
                var btnCol = new DataGridViewButtonColumn
                {
                    Name = "Select",
                    HeaderText = "",
                    Text = "📁",
                    UseColumnTextForButtonValue = true,
                    Width = (int)(40 * scale)
                };
                var folderCol = new DataGridViewTextBoxColumn
                {
                    Name = "Folder",
                    HeaderText = sFolder,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                    ReadOnly = false
                };

                grid.ColumnHeadersHeight = (int)(32 * scale);

                grid.Columns.Add(monitorCol);
                grid.Columns.Add(timePickerCol);
                grid.Columns.Add(waitCol);
                grid.Columns.Add(btnCol);
                grid.Columns.Add(folderCol);

                // Disable sorting on all columns
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }

                grid.RowTemplate.Height = (int)(32 * scale);

                // Create one row per monitor, showing monitor number starting at 1
                for (int i = 0; i < monitorCount; i++)
                {
                    string folder = (i < folderWaits.Count) ? folderWaits[i].folder : "";
                    int wait = (i < folderWaits.Count) ? folderWaits[i].wait : 60;
                    grid.Rows.Add((i + 1).ToString(), null, wait.ToString(), null, folder);
                }

                int rowHeight = grid.Rows.Count > 0 ? grid.Rows[0].Height : (int)(grid.RowTemplate.Height * scale);
                int headerHeight = grid.ColumnHeadersHeight;
                int gridHeight = grid.Rows.Count * rowHeight + headerHeight;
                int buttonWidth = (int)(90 * scale);
                int buttonHeight = (int)(30 * scale);
                int padding = (int)(16 * scale);

                grid.Left = 0;
                grid.Top = 0;
                grid.Width = this.ClientSize.Width - 1;
                grid.Height = gridHeight;

                btnOK = new Button
                {
                    Text = sOK,
                    Width = buttonWidth,
                    Height = buttonHeight,
                    Left = padding,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                };
                btnHelp = new Button
                {
                    Text = sHelp,
                    Width = buttonWidth,
                    Height = buttonHeight,
                    Left = btnOK.Left + buttonWidth + padding,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                };
                btnExit = new Button
                {
                    Text = sExit,
                    Width = buttonWidth,
                    Height = buttonHeight,
                    Left = btnHelp.Left + buttonWidth + padding,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                };

                bool isDark = Dark;
                if (isDark)
                {
                    // Apply dark mode styling to form buttons
                    foreach (var btn in new[] { btnOK, btnHelp, btnExit })
                    {
                        btn.FlatStyle = FlatStyle.Flat;
                        btn.FlatAppearance.BorderColor = SystemColors.Highlight;
                        btn.FlatAppearance.BorderSize = 1;
                        btn.BackColor = Color.FromArgb(60, 60, 60);
                        btn.FlatAppearance.MouseOverBackColor = Color.Black;
                        btn.ForeColor = Color.White;
                    }

                    // Apply dark mode styling to form and grid
                    this.BackColor = Color.FromArgb(32, 32, 32);
                    grid.BackgroundColor = Color.FromArgb(32, 32, 32);
                    grid.DefaultCellStyle.BackColor = Color.FromArgb(32, 32, 32);
                    grid.DefaultCellStyle.ForeColor = Color.White;
                    grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(48, 48, 48);
                    grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    grid.EnableHeadersVisualStyles = false;
                    grid.GridColor = Color.Black;

                    grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(48, 48, 48);
                    grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.White;
                    grid.ColumnHeadersDefaultCellStyle.Padding = new Padding(0, 0, 0, 0);
                    DarkTitleBar(this.Handle);

                    grid.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;

                    // Make Monitor column selection invisible in dark mode
                    monitorCol.DefaultCellStyle.SelectionBackColor = Color.FromArgb(32, 32, 32);
                    monitorCol.DefaultCellStyle.SelectionForeColor = Color.White;
                }
                else
                {
                    // Make Monitor column selection invisible in light mode
                    monitorCol.DefaultCellStyle.SelectionBackColor = SystemColors.Window;
                    monitorCol.DefaultCellStyle.SelectionForeColor = SystemColors.WindowText;

                    // Disable hover effect on column headers in light mode
                    grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = grid.ColumnHeadersDefaultCellStyle.BackColor;
                }

                btnOK.Click += BtnOK_Click;
                btnHelp.Click += BtnHelp_Click;
                btnExit.Click += BtnExit_Click;

                // Enable custom painting for button cells in both light and dark mode
                grid.CellPainting += Grid_CellPainting;
                grid.CellMouseMove += Grid_CellMouseMove;
                grid.CellMouseLeave += Grid_CellMouseLeave;

                int spaceAboveButtons = (int)(12 * scale);
                int spaceBelowButtons = (int)(8 * scale);

                int buttonsTop = grid.Top + grid.Height + spaceAboveButtons;
                this.ClientSize = new Size(
                    this.Width - padding,
                    buttonsTop + buttonHeight + spaceBelowButtons
                );

                btnOK.Top = buttonsTop;
                btnHelp.Top = buttonsTop;
                btnExit.Top = buttonsTop;

                toggleStartup = new ToggleSwitch
                {
                    Checked = IsStartup,
                    Left = btnExit.Left + buttonWidth + padding,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Left
                };
                toggleStartup.Top = buttonsTop + (buttonHeight - toggleStartup.Height) / 2;

                lblStartup = new Label
                {
                    Text = sStartup,
                    AutoSize = true,
                    Left = toggleStartup.Left + toggleStartup.Width + (int)(8 * scale),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Left
                };
                lblStartup.Top = buttonsTop + (buttonHeight - lblStartup.Height) / 2;

                if (isDark) lblStartup.ForeColor = Color.White;

                this.Controls.Add(grid);
                this.Controls.Add(btnOK);
                this.Controls.Add(btnHelp);
                this.Controls.Add(btnExit);
                this.Controls.Add(toggleStartup);
                this.Controls.Add(lblStartup);

                grid.CellContentClick += Grid_CellContentClick;
                grid.EditingControlShowing += Grid_EditingControlShowing;
                toggleStartup.CheckedChanged += ToggleStartup_CheckedChanged;

                this.StartPosition = FormStartPosition.CenterScreen;
            }

            // Trigger repaint when mouse moves over button cells for hover effect
            private void Grid_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
            {
                if (e.RowIndex >= 0 && (e.ColumnIndex == grid.Columns["Select"].Index || e.ColumnIndex == grid.Columns["TimePicker"].Index))
                {
                    if (hoveredButtonRow != e.RowIndex || hoveredButtonCol != e.ColumnIndex)
                    {
                        // Invalidate previous hovered cell
                        if (hoveredButtonRow >= 0 && hoveredButtonCol >= 0)
                        {
                            grid.InvalidateCell(hoveredButtonCol, hoveredButtonRow);
                        }
                        
                        // Update and invalidate new hovered cell
                        hoveredButtonRow = e.RowIndex;
                        hoveredButtonCol = e.ColumnIndex;
                        grid.InvalidateCell(e.ColumnIndex, e.RowIndex);
                    }
                }
            }

            private void Grid_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex >= 0 && (e.ColumnIndex == grid.Columns["Select"].Index || e.ColumnIndex == grid.Columns["TimePicker"].Index))
                {
                    hoveredButtonRow = -1;
                    hoveredButtonCol = -1;
                    grid.InvalidateCell(e.ColumnIndex, e.RowIndex);
                }
            }

            // Custom paint the button cells (⏱ and 📁)
            private void Grid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
            {
                if ((e.ColumnIndex == grid.Columns["Select"].Index || e.ColumnIndex == grid.Columns["TimePicker"].Index) && e.RowIndex >= 0)
                {
                    bool isDark = Dark;
                    
                    // Paint background
                    e.PaintBackground(e.CellBounds, false);

                    // Define button area with padding
                    Rectangle buttonRect = e.CellBounds;
                    buttonRect.Inflate(-2, -2);

                    // Determine if button is being hovered
                    bool isHovered = (hoveredButtonRow == e.RowIndex && hoveredButtonCol == e.ColumnIndex);

                    // Draw button background
                    Color buttonColor = isHovered 
                        ? (isDark ? Color.Black : Color.FromArgb(229, 241, 251))
                        : (isDark ? Color.FromArgb(60, 60, 60) : SystemColors.Control);
                    using (SolidBrush brush = new SolidBrush(buttonColor))
                    {
                        e.Graphics.FillRectangle(brush, buttonRect);
                    }

                    // Draw button border in light mode
                    if (!isDark)
                    {
                        ControlPaint.DrawBorder(e.Graphics, buttonRect, SystemColors.ControlDark, ButtonBorderStyle.Solid);
                    }

                    // Draw button text with larger font using Segoe UI Emoji for proper emoji support
                    string text = e.ColumnIndex == grid.Columns["Select"].Index ? "📁" : "⏱";
                    using (Font emojiFont = new Font("Segoe UI Emoji", e.CellStyle.Font.Size + 1, FontStyle.Regular))
                    {
                        using (StringFormat sf = new StringFormat())
                        {
                            sf.Alignment = StringAlignment.Center;
                            sf.LineAlignment = StringAlignment.Center;
                            Color textColor = isDark ? Color.White : SystemColors.ControlText;
                            using (SolidBrush textBrush = new SolidBrush(textColor))
                            {
                                e.Graphics.DrawString(text, emojiFont, textBrush, buttonRect, sf);
                            }
                        }
                    }

                    e.Handled = true;
                }
            }

            // Validate that all folders are set and exist before saving
            private void BtnOK_Click(object sender, EventArgs e)
            {
                // Validate all folder fields
                for (int i = 0; i < grid.Rows.Count; i++)
                {
                    string folder = grid.Rows[i].Cells["Folder"].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(folder))
                    {
                        grid.CurrentCell = grid.Rows[i].Cells["Folder"];
                        return;
                    }
                    if (!System.IO.Directory.Exists(folder))
                    {
                        grid.CurrentCell = grid.Rows[i].Cells["Folder"];
                        return;
                    }
                }

                SaveSettings();
                StartWallpaperSlideshowIfNotRunning();
                this.Close();
            }

            private void BtnHelp_Click(object sender, EventArgs e)
            {
                Process.Start("https://lesferch.github.io/WallpaperSlideshow/");
            }

            // Exit button stops the slideshow and closes the UI
            private void BtnExit_Click(object sender, EventArgs e)
            {
                // Kill all WallpaperSlideshow instances
                KillAllWallpaperSlideshowInstances();

                // Signal the main app to exit
                try
                {
                    using (var exitEvent = new System.Threading.EventWaitHandle(false, System.Threading.EventResetMode.ManualReset, "WallpaperSlideshowExit"))
                    {
                        exitEvent.Set();
                    }
                }
                catch { }
                Application.Exit();
            }

            // Save settings to registry: HKCU\Software\WallpaperSlideshow\0, \1, \2, etc.
            private void SaveSettings()
            {
                using (var baseKey = Registry.CurrentUser.CreateSubKey(@"Software\WallpaperSlideshow"))
                {
                    for (int i = 0; i < grid.Rows.Count; i++)
                    {
                        string folder = grid.Rows[i].Cells["Folder"].Value?.ToString() ?? "";
                        int wait = 60;
                        int.TryParse(grid.Rows[i].Cells["Wait"].Value?.ToString(), out wait);
                        if (wait <= 0) wait = 60;
                        if (wait > MAX_WAIT_SECONDS) wait = MAX_WAIT_SECONDS; // Cap at maximum
                        using (var subKey = baseKey.CreateSubKey(i.ToString()))
                        {
                            subKey.SetValue("", folder, RegistryValueKind.String);
                            subKey.SetValue("Wait", wait, RegistryValueKind.DWord);
                        }
                    }
                }
            }

            private void StartWallpaperSlideshowIfNotRunning()
            {
                var processes = Process.GetProcessesByName("WallpaperSlideshow");
                bool isRunning = processes.Length > 0;
                foreach (var process in processes)
                {
                    process.Dispose();
                }

                if (!isRunning)
                {
                    try
                    {
                        string exePath = System.IO.Path.Combine(
                            System.IO.Path.GetDirectoryName(Application.ExecutablePath),
                            "WallpaperSlideshow.exe");

                        if (System.IO.File.Exists(exePath))
                        {
                            Process.Start(exePath);
                        }
                    }
                    catch
                    {
                        // Failed to start process
                    }
                }
            }

            private void KillAllWallpaperSlideshowInstances()
            {
                var processes = Process.GetProcessesByName("WallpaperSlideshow");

                foreach (var process in processes)
                {
                    try
                    {
                        process.Kill();
                    }
                    catch
                    {
                        // Process may have already exited or access denied
                    }
                }

                // Clean up process objects
                foreach (var process in processes)
                {
                    process.Dispose();
                }
            }

            // Handle clicks on button cells: ⏱ opens time picker, 📁 opens folder picker
            private void Grid_CellContentClick(object sender, DataGridViewCellEventArgs e)
            {
                if (e.ColumnIndex == grid.Columns["Select"].Index && e.RowIndex >= 0)
                {
                    string currentFolder = grid.Rows[e.RowIndex].Cells["Folder"].Value?.ToString();

                    FolderPicker fd = new FolderPicker
                    {
                        Title = "", // This will display "Select Folder" in the current OS language
                        InputPath = !string.IsNullOrEmpty(currentFolder) && System.IO.Directory.Exists(currentFolder) ? currentFolder : "",
                        Multiselect = false
                    };
                    if (fd.ShowDialog(IntPtr.Zero) == true && !string.IsNullOrEmpty(fd.ResultPath))
                    {
                        grid.Rows[e.RowIndex].Cells["Folder"].Value = fd.ResultPath;
                    }
                }
                else if (e.ColumnIndex == grid.Columns["TimePicker"].Index && e.RowIndex >= 0)
                {
                    int currentSeconds = 60;
                    int.TryParse(grid.Rows[e.RowIndex].Cells["Wait"].Value?.ToString(), out currentSeconds);
                    
                    // Cap at maximum time picker value
                    if (currentSeconds > MAX_WAIT_SECONDS)
                    {
                        currentSeconds = MAX_WAIT_SECONDS;
                        grid.Rows[e.RowIndex].Cells["Wait"].Value = currentSeconds.ToString();
                    }

                    using (var timePicker = new TimePickerDialog(currentSeconds))
                    {
                        if (timePicker.ShowDialog(this) == DialogResult.OK)
                        {
                            grid.Rows[e.RowIndex].Cells["Wait"].Value = timePicker.TotalSeconds.ToString();
                        }
                    }
                }
            }

            // Set editing font and restrict Wait column to numeric input only
            private void Grid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
            {
                var tb = e.Control as TextBox;
                if (tb != null)
                {
                    // Set editing font to match cell font
                    tb.Font = grid.CurrentCell.InheritedStyle.Font;

                    // Add numeric validation only for Wait column
                    if (grid.CurrentCell.ColumnIndex == grid.Columns["Wait"].Index)
                    {
                        tb.KeyPress -= WaitColumn_KeyPress;
                        tb.KeyPress += WaitColumn_KeyPress;
                    }
                    else
                    {
                        tb.KeyPress -= WaitColumn_KeyPress;
                    }
                }
            }

            private void WaitColumn_KeyPress(object sender, KeyPressEventArgs e)
            {
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
                {
                    e.Handled = true;
                }
            }

            // Handle startup toggle changes
            private void ToggleStartup_CheckedChanged(object sender, EventArgs e)
            {
                if (toggleStartup.Checked)
                {
                    CreateStartupShortcut();
                    IsStartup = true;
                }
                else
                {
                    RemoveStartupShortcut();
                    IsStartup = false;
                }
            }
        }

        // Courtesy of Simon Mourier https://stackoverflow.com/a/66187224/15764378
        public class FolderPicker
        {
            private readonly List<string> _resultPaths = new List<string>();
            private readonly List<string> _resultNames = new List<string>();

            public IReadOnlyList<string> ResultPaths => _resultPaths;
            public IReadOnlyList<string> ResultNames => _resultNames;
            public string ResultPath => ResultPaths.FirstOrDefault();
            public string ResultName => ResultNames.FirstOrDefault();
            public virtual string InputPath { get; set; }
            public virtual bool ForceFileSystem { get; set; }
            public virtual bool Multiselect { get; set; }
            public virtual string Title { get; set; }
            public virtual string OkButtonLabel { get; set; }
            public virtual string FileNameLabel { get; set; }
            protected virtual int SetOptions(int options)
            {
                if (ForceFileSystem)
                {
                    options |= (int)FOS.FOS_FORCEFILESYSTEM;
                }

                if (Multiselect)
                {
                    options |= (int)FOS.FOS_ALLOWMULTISELECT;
                }
                return options;
            }
            public virtual bool? ShowDialog(IntPtr owner, bool throwOnError = false)
            {
                var dialog = (IFileOpenDialog)new FileOpenDialog();
                if (!string.IsNullOrEmpty(InputPath))
                {
                    if (CheckHr(SHCreateItemFromParsingName(InputPath, null, typeof(IShellItem).GUID, out var item), throwOnError) != 0)
                        return null;
                    dialog.SetFolder(item);
                }
                var options = FOS.FOS_PICKFOLDERS;
                options = (FOS)SetOptions((int)options);
                dialog.SetOptions(options);
                if (Title != null)
                {
                    dialog.SetTitle(Title);
                }
                if (OkButtonLabel != null)
                {
                    dialog.SetOkButtonLabel(OkButtonLabel);
                }
                if (FileNameLabel != null)
                {
                    dialog.SetFileNameLabel(FileNameLabel);
                }
                if (owner == IntPtr.Zero)
                {
                    owner = Process.GetCurrentProcess().MainWindowHandle;
                    if (owner == IntPtr.Zero)
                    {
                        owner = GetDesktopWindow();
                    }
                }
                var hr = dialog.Show(owner);
                if (hr == ERROR_CANCELLED)
                    return null;
                if (CheckHr(hr, throwOnError) != 0)
                    return null;

                if (CheckHr(dialog.GetResults(out var items), throwOnError) != 0)
                    return null;

                items.GetCount(out var count);
                for (var i = 0; i < count; i++)
                {
                    items.GetItemAt(i, out var item);
                    CheckHr(item.GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEPARSING, out var path), throwOnError);
                    CheckHr(item.GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEEDITING, out var name), throwOnError);
                    if (path != null || name != null)
                    {
                        _resultPaths.Add(path);
                        _resultNames.Add(name);
                    }
                }
                return true;
            }
            private static int CheckHr(int hr, bool throwOnError)
            {
                if (hr != 0 && throwOnError) Marshal.ThrowExceptionForHR(hr);
                return hr;
            }
            [DllImport("shell32")]
            private static extern int SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string pszPath, IBindCtx pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);
            [DllImport("user32")]
            private static extern IntPtr GetDesktopWindow();
            private const int ERROR_CANCELLED = unchecked((int)0x800704C7);
            [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")] // CLSID_FileOpenDialog
            private class FileOpenDialog { }

            [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            private interface IFileOpenDialog
            {
                [PreserveSig] int Show(IntPtr parent); // IModalWindow
                [PreserveSig] int SetFileTypes();  // not fully defined
                [PreserveSig] int SetFileTypeIndex(int iFileType);
                [PreserveSig] int GetFileTypeIndex(out int piFileType);
                [PreserveSig] int Advise(); // not fully defined
                [PreserveSig] int Unadvise();
                [PreserveSig] int SetOptions(FOS fos);
                [PreserveSig] int GetOptions(out FOS pfos);
                [PreserveSig] int SetDefaultFolder(IShellItem psi);
                [PreserveSig] int SetFolder(IShellItem psi);
                [PreserveSig] int GetFolder(out IShellItem ppsi);
                [PreserveSig] int GetCurrentSelection(out IShellItem ppsi);
                [PreserveSig] int SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
                [PreserveSig] int GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
                [PreserveSig] int SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
                [PreserveSig] int SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
                [PreserveSig] int SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
                [PreserveSig] int GetResult(out IShellItem ppsi);
                [PreserveSig] int AddPlace(IShellItem psi, int alignment);
                [PreserveSig] int SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
                [PreserveSig] int Close(int hr);
                [PreserveSig] int SetClientGuid();  // not fully defined
                [PreserveSig] int ClearClientData();
                [PreserveSig] int SetFilter([MarshalAs(UnmanagedType.IUnknown)] object pFilter);
                [PreserveSig] int GetResults(out IShellItemArray ppenum);
                [PreserveSig] int GetSelectedItems([MarshalAs(UnmanagedType.IUnknown)] out object ppsai);
            }
            [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            private interface IShellItem
            {
                [PreserveSig] int BindToHandler(); // not fully defined
                [PreserveSig] int GetParent(); // not fully defined
                [PreserveSig] int GetDisplayName(SIGDN sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
                [PreserveSig] int GetAttributes();  // not fully defined
                [PreserveSig] int Compare();  // not fully defined
            }

            [ComImport, Guid("b63ea76d-1f85-456f-a19c-48159efa858b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            private interface IShellItemArray
            {
                [PreserveSig] int BindToHandler();  // not fully defined
                [PreserveSig] int GetPropertyStore();  // not fully defined
                [PreserveSig] int GetPropertyDescriptionList();  // not fully defined
                [PreserveSig] int GetAttributes();  // not fully defined
                [PreserveSig] int GetCount(out int pdwNumItems);
                [PreserveSig] int GetItemAt(int dwIndex, out IShellItem ppsi);
                [PreserveSig] int EnumItems();  // not fully defined
            }

            private enum SIGDN : uint
            {
                SIGDN_DESKTOPABSOLUTEEDITING = 0x8004c000,
                SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000,
                SIGDN_FILESYSPATH = 0x80058000,
                SIGDN_NORMALDISPLAY = 0,
                SIGDN_PARENTRELATIVE = 0x80080001,
                SIGDN_PARENTRELATIVEEDITING = 0x80031001,
                SIGDN_PARENTRELATIVEFORADDRESSBAR = 0x8007c001,
                SIGDN_PARENTRELATIVEPARSING = 0x80018001,
                SIGDN_URL = 0x80068000
            }
            [Flags]
            private enum FOS
            {
                FOS_OVERWRITEPROMPT = 0x2,
                FOS_STRICTFILETYPES = 0x4,
                FOS_NOCHANGEDIR = 0x8,
                FOS_PICKFOLDERS = 0x20,
                FOS_FORCEFILESYSTEM = 0x40,
                FOS_ALLNONSTORAGEITEMS = 0x80,
                FOS_NOVALIDATE = 0x100,
                FOS_ALLOWMULTISELECT = 0x200,
                FOS_PATHMUSTEXIST = 0x800,
                FOS_FILEMUSTEXIST = 0x1000,
                FOS_CREATEPROMPT = 0x2000,
                FOS_SHAREAWARE = 0x4000,
                FOS_NOREADONLYRETURN = 0x8000,
                FOS_NOTESTFILECREATE = 0x10000,
                FOS_HIDEMRUPLACES = 0x20000,
                FOS_HIDEPINNEDPLACES = 0x40000,
                FOS_NODEREFERENCELINKS = 0x100000,
                FOS_OKBUTTONNEEDSINTERACTION = 0x200000,
                FOS_DONTADDTORECENT = 0x2000000,
                FOS_FORCESHOWHIDDEN = 0x10000000,
                FOS_DEFAULTNOMINIMODE = 0x20000000,
                FOS_FORCEPREVIEWPANEON = 0x40000000,
                FOS_SUPPORTSTREAMABLEITEMS = unchecked((int)0x80000000)
            }
        }

        // Custom toggle switch control similar to Windows 10 Settings
        public class ToggleSwitch : Control
        {
            private bool _checked = false;
            private float _animationProgress = 0;
            private System.Windows.Forms.Timer _animationTimer;
            private const int AnimationSteps = 10;
            private bool _isHovered = false;

            public bool Checked
            {
                get { return _checked; }
                set
                {
                    if (_checked != value)
                    {
                        _checked = value;
                        StartAnimation();
                        CheckedChanged?.Invoke(this, EventArgs.Empty);
                    }
                }
            }

            public event EventHandler CheckedChanged;

            public ToggleSwitch()
            {
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                this.Size = new Size(44, 22);

                _animationTimer = new System.Windows.Forms.Timer();
                _animationTimer.Interval = 15;
                _animationTimer.Tick += AnimationTimer_Tick;
            }

            private void StartAnimation()
            {
                _animationTimer.Start();
            }

            private void AnimationTimer_Tick(object sender, EventArgs e)
            {
                if (_checked)
                {
                    _animationProgress += 1.0f / AnimationSteps;
                    if (_animationProgress >= 1.0f)
                    {
                        _animationProgress = 1.0f;
                        _animationTimer.Stop();
                    }
                }
                else
                {
                    _animationProgress -= 1.0f / AnimationSteps;
                    if (_animationProgress <= 0.0f)
                    {
                        _animationProgress = 0.0f;
                        _animationTimer.Stop();
                    }
                }
                this.Invalidate();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                bool isDark = Dark;

                // Different hover colors for light and dark mode
                Color hoverColor = isDark ? Color.Black : Color.FromArgb(77, 161, 227);

                // Windows 10 style colors
                Color trackColorOff = _isHovered
                    ? hoverColor
                    : (isDark ? Color.FromArgb(60, 60, 60) : Color.FromArgb(200, 200, 200));

                Color trackColorOn = _isHovered
                    ? hoverColor
                    : Color.FromArgb(0, 120, 215); // Windows 10 blue

                // Different thumb colors for light and dark mode
                Color thumbColor = Color.White;
                Color borderColor = isDark ? Color.FromArgb(80, 80, 80) : Color.FromArgb(180, 180, 180);

                // Interpolate colors based on animation progress
                int r = (int)(trackColorOff.R + (trackColorOn.R - trackColorOff.R) * _animationProgress);
                int g = (int)(trackColorOff.G + (trackColorOn.G - trackColorOff.G) * _animationProgress);
                int b = (int)(trackColorOff.B + (trackColorOn.B - trackColorOff.B) * _animationProgress);
                Color currentTrackColor = Color.FromArgb(r, g, b);

                // Draw track (rounded rectangle) - inset by 1 pixel to account for border
                int trackHeight = this.Height - 2;  // Preserve space for border
                int trackWidth = this.Width - 2;
                Rectangle trackRect = new Rectangle(1, 1, trackWidth, trackHeight);

                using (System.Drawing.Drawing2D.GraphicsPath trackPath = GetRoundedRectangle(trackRect, trackHeight / 2))
                {
                    using (SolidBrush trackBrush = new SolidBrush(currentTrackColor))
                    {
                        e.Graphics.FillPath(trackBrush, trackPath);
                    }

                    // Draw border
                    using (Pen borderPen = new Pen(borderColor, 1))
                    {
                        e.Graphics.DrawPath(borderPen, trackPath);
                    }
                }

                // Calculate thumb position - smaller thumb for Windows 10 style with equal padding on both sides
                int thumbSize = trackHeight - 8; // Smaller by 4 pixels for more visible background
                int thumbMaxX = trackWidth - thumbSize - 8; // 4 pixels padding on each side
                int thumbX = (int)(4 + thumbMaxX * _animationProgress);
                int thumbY = 4;

                // Draw thumb (circle with shadow)
                Rectangle thumbRect = new Rectangle(thumbX, thumbY, thumbSize, thumbSize);

                // Shadow
                Rectangle shadowRect = new Rectangle(thumbX + 1, thumbY + 1, thumbSize, thumbSize);
                using (System.Drawing.Drawing2D.GraphicsPath shadowPath = new System.Drawing.Drawing2D.GraphicsPath())
                {
                    shadowPath.AddEllipse(shadowRect);
                    using (System.Drawing.Drawing2D.PathGradientBrush shadowBrush =
                           new System.Drawing.Drawing2D.PathGradientBrush(shadowPath))
                    {
                        shadowBrush.CenterColor = Color.FromArgb(30, 0, 0, 0);
                        shadowBrush.SurroundColors = new[] { Color.FromArgb(0, 0, 0, 0) };
                        e.Graphics.FillEllipse(shadowBrush, shadowRect);
                    }
                }

                // Thumb
                using (SolidBrush thumbBrush = new SolidBrush(thumbColor))
                {
                    e.Graphics.FillEllipse(thumbBrush, thumbRect);
                }
            }

            private System.Drawing.Drawing2D.GraphicsPath GetRoundedRectangle(Rectangle bounds, int radius)
            {
                int diameter = radius * 2;
                Size size = new Size(diameter, diameter);
                Rectangle arc = new Rectangle(bounds.Location, size);
                System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();

                if (radius == 0)
                {
                    path.AddRectangle(bounds);
                    return path;
                }

                // Top left arc
                path.AddArc(arc, 180, 90);

                // Top right arc
                arc.X = bounds.Right - diameter;
                path.AddArc(arc, 270, 90);

                // Bottom right arc
                arc.Y = bounds.Bottom - diameter;
                path.AddArc(arc, 0, 90);

                // Bottom left arc
                arc.X = bounds.Left;
                path.AddArc(arc, 90, 90);

                path.CloseFigure();
                return path;
            }

            protected override void OnClick(EventArgs e)
            {
                base.OnClick(e);
                Checked = !Checked;
            }

            protected override void OnMouseEnter(EventArgs e)
            {
                base.OnMouseEnter(e);
                _isHovered = true;
                this.Invalidate();
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                base.OnMouseLeave(e);
                _isHovered = false;
                this.Invalidate();
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _animationTimer?.Stop();
                    _animationTimer?.Dispose();
                }
                base.Dispose(disposing);
            }
        }
    }
}
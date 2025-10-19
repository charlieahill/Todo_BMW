using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Threading;
using System.Windows.Interop;
using WinForms = System.Windows.Forms;

namespace Todo
{
    public partial class MainWindow : Window
    {
        // Storage paths
        private static readonly string DataDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "CHillSW", "TodoBMW");
        private static readonly string SaveFileName = Path.Combine(DataDirectory, "tasks_by_date.json");
        private static readonly string MetaFileName = Path.Combine(DataDirectory, "meta.json");
        private static readonly string StartupLogFile = Path.Combine(DataDirectory, "startup_checks.log");
        private static readonly string DebugLogFile = Path.Combine(DataDirectory, "tasks_debug_log.txt");

        private const string RepositoryKey = "__repository__";

        private Dictionary<string, List<TaskModel>> AllTasks { get; set; } = new Dictionary<string, List<TaskModel>>();
        public ObservableCollection<TaskModel> TaskList { get; set; } = new ObservableCollection<TaskModel>();

        private DateTime _currentDate = DateTime.Today;
        private bool _isInitializing = false;

        // Auto-backup
        private DispatcherTimer _autoBackupTimer;
        private static readonly string AutoBackupFolder = Path.Combine(DataDirectory, "autobackups");
        private const int AutoBackupKeep = 20;
        private readonly TimeSpan AutoBackupInterval = TimeSpan.FromHours(1);

        // Dock/snap tracking
        private double _initialDockedHeight = 0; // working-area height at last dock
        private bool _isDockingOperation = false; // suppress LocationChanged while we move programmatically
        private bool _hasShrunkAfterMove = false; // ensure half-height is applied only once after docking
        private DispatcherTimer _shrinkAfterMoveTimer; // debounce applying half-height until after move settles
        private double? _targetShrinkHeight = null; // value to enforce briefly against OS snap override
        private int _shrinkEnforcementAttempts = 0;

        // Track move/size session to detect start and end of manual move
        private bool _isMoveOrSize = false;
        private bool _wasRightFullHeightAtMoveStart = false;

        // misc flags
        private bool _suppressTextChanged = false;
        private bool _placeholderJustFocused = false;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            _isInitializing = true;

            try { Directory.CreateDirectory(DataDirectory); } catch { }

            try
            {
                try { File.WriteAllText(StartupLogFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Startup log cleared\n"); } catch { }
                try { File.WriteAllText(DebugLogFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Debug log cleared\n"); } catch { }
            }
            catch { }

            LoadTasks();
            SetCurrentDate(_currentDate);

            // If application was last opened on a previous day, prompt to carry over unfinished tasks
            // Moved carryover handling to Loaded event so dialog is shown reliably after window is displayed
            this.Loaded += MainWindow_Loaded;

            // Track manual moves (now no-op; move end handled via WM_EXITSIZEMOVE)
            this.LocationChanged += MainWindow_LocationChanged;
            // this.SizeChanged += MainWindow_SizeChanged; // removed: we only shrink at end of move

            // Removed debounce timer; shrink occurs at end of move via WM_EXITSIZEMOVE
            // _shrinkAfterMoveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
            // _shrinkAfterMoveTimer.Tick += (s, e) => { try { _shrinkAfterMoveTimer.Stop(); PerformHalfHeightShrinkIfNeeded(); } catch { } };

            lbTasksList.ItemsSource = TaskList;
            lbTasksList.MouseDoubleClick += LbTasksList_MouseDoubleClick;

            Application.Current.Exit += Current_Exit;

            // Start automatic backups
            try
            {
                _autoBackupTimer = new DispatcherTimer { Interval = AutoBackupInterval };
                _autoBackupTimer.Tick += (s, e) => { try { AutoBackupNow(); } catch { } };
                _autoBackupTimer.Start();
            }
            catch { }

            // Initialize last backup display
            try
            {
                var latest = FindLatestBackupTimestamp();
                if (latest.HasValue)
                    UpdateLastBackupText(latest.Value);
                else
                    UpdateLastBackupText(null);
            }
            catch { UpdateLastBackupText(null); }

            // Initialize last saved display (based on file timestamp if present)
            try
            {
                if (File.Exists(SaveFileName))
                    UpdateLastSavedText(File.GetLastWriteTime(SaveFileName));
                else
                    UpdateLastSavedText(null);
            }
            catch { UpdateLastSavedText(null); }

            // Initialize last opened (previous run) display if meta exists
            try
            {
                if (File.Exists(MetaFileName))
                {
                    var jm = File.ReadAllText(MetaFileName);
                    var meta = JsonSerializer.Deserialize<MetaInfo>(jm);
                    UpdateLastOpenedText(meta?.LastOpened);
                }
                else
                {
                    UpdateLastOpenedText(null);
                }
            }
            catch { UpdateLastOpenedText(null); }

            // Subscribe to TaskList changes to update view counts
            TaskList.CollectionChanged += (s, e) => UpdateTotals();

            _isInitializing = false;

            // Record open event
            try { TimeTrackingService.Instance.RecordOpen(); } catch { }

            // Update totals at startup
            UpdateTotals();

            // Ensure Today search UI visibility for initial mode
            try { if (TodaySearchContainer != null) TodaySearchContainer.Visibility = Visibility.Visible; } catch { }

            // Set initial Tag for mode-based triggers
            this.Tag = "Today";
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Dock to right edge full height by default on first show
                DockRightFullHeight();

                // Ensure carryover dialog is invoked after the main window is shown so it appears reliably
                Dispatcher.BeginInvoke(new Action(() => HandleCarryOverIfNewDay()), DispatcherPriority.Background);
            }
            catch { }
            finally
            {
                // Unsubscribe to avoid re-invoking if Loaded fires again
                this.Loaded -= MainWindow_Loaded;
            }
        }

        // DockRightFullHeight modifies _initialDockedHeight, which is used to decide half-height on manual move.
        private void MainWindow_LocationChanged(object sender, EventArgs e)
        {
            try
            {
                // no-op: shrink handled in WM_EXITSIZEMOVE
            }
            catch { }
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            try
            {
                var hwnd = new WindowInteropHelper(this).Handle;
                var source = HwndSource.FromHwnd(hwnd);
                if (source != null)
                {
                    source.AddHook(WndProc);
                }
            }
            catch { }
        }

        private const int WM_ENTERSIZEMOVE = 0x0231;
        private const int WM_EXITSIZEMOVE = 0x0232;

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            try
            {
                switch (msg)
                {
                    case WM_ENTERSIZEMOVE:
                        _isMoveOrSize = true;
                        _wasRightFullHeightAtMoveStart = IsAtRightFullHeight();
                        // Ensure the baseline height uses current working area when at full height
                        if (_wasRightFullHeightAtMoveStart)
                        {
                            _initialDockedHeight = GetWorkingAreaDips().height;
                        }
                        break;
                    case WM_EXITSIZEMOVE:
                        _isMoveOrSize = false;
                        if (_wasRightFullHeightAtMoveStart)
                        {
                            ApplyHalfHeightFromWorkingArea();
                            _hasShrunkAfterMove = true;
                        }
                        _wasRightFullHeightAtMoveStart = false;
                        break;
                }
            }
            catch { }
            return IntPtr.Zero;
        }

        private (double left, double top, double width, double height) GetWorkingAreaDips()
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            var screen = WinForms.Screen.FromHandle(hwnd);
            var wa = screen.WorkingArea; // pixels
            var dpi = VisualTreeHelper.GetDpi(this);
            return (wa.Left / dpi.DpiScaleX,
                    wa.Top / dpi.DpiScaleY,
                    wa.Width / dpi.DpiScaleX,
                    wa.Height / dpi.DpiScaleY);
        }

        private bool IsAtRightFullHeight()
        {
            if (WindowState != WindowState.Normal) return false;
            var wa = GetWorkingAreaDips();
            double epsilon = 2.0;
            double right = Left + Width;
            double waRight = wa.left + wa.width;
            return Math.Abs(Top - wa.top) <= epsilon
                   && Math.Abs(Height - wa.height) <= epsilon
                   && Math.Abs(right - waRight) <= epsilon;
        }

        private void ApplyHalfHeightFromWorkingArea()
        {
            try
            {
                var wa = GetWorkingAreaDips();
                var newHeight = wa.height / 2.0;
                if (MinHeight > 0 && newHeight < MinHeight) newHeight = MinHeight;
                _isDockingOperation = true;
                Height = newHeight;
            }
            catch { }
            finally
            {
                Dispatcher.BeginInvoke(new Action(() => _isDockingOperation = false), DispatcherPriority.Background);
            }
        }

        private void DockRightFullHeightButton_Click(object sender, RoutedEventArgs e)
        {
            DockRightFullHeight();
        }

        private void DockRightFullHeight()
        {
            try
            {
                _isDockingOperation = true;

                // Restore from Maximized to allow manual sizing
                if (this.WindowState != WindowState.Normal)
                    this.WindowState = WindowState.Normal;

                var hwnd = new WindowInteropHelper(this).Handle;
                var screen = WinForms.Screen.FromHandle(hwnd);
                var wa = screen.WorkingArea; // pixels

                // Get DPI for this window/monitor to convert to WPF DIPs
                var dpi = VisualTreeHelper.GetDpi(this);
                double dipLeft = wa.Left / dpi.DpiScaleX;
                double dipTop = wa.Top / dpi.DpiScaleY;
                double dipWidth = wa.Width / dpi.DpiScaleX;
                double dipHeight = wa.Height / dpi.DpiScaleY;

                // Desired width: keep current width but clamp to work area
                double width = double.IsNaN(this.Width) || this.Width <= 0 ? (this.ActualWidth > 0 ? this.ActualWidth : dipWidth) : this.Width;
                if (width > dipWidth) width = dipWidth;

                this.Top = dipTop;
                this.Height = dipHeight;
                this.Left = dipLeft + (dipWidth - width);
                this.Width = width;

                _initialDockedHeight = dipHeight;
                _hasShrunkAfterMove = false; // allow the next manual move to trigger half-height

                // stop any pending shrink requests scheduled prior to docking again
                _shrinkAfterMoveTimer?.Stop();
            }
            catch { }
            finally
            {
                // Clear docking flag shortly after layout updates to ignore LocationChanged during programmatic move
                Dispatcher.BeginInvoke(new Action(() => _isDockingOperation = false), DispatcherPriority.Background);
            }
        }

        private void AppendStartupLog(string message)
        {
            try
            {
                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n";
                // Ensure data directory exists
                try { Directory.CreateDirectory(DataDirectory); } catch { }
                File.AppendAllText(StartupLogFile, line);
                // Also append to the general debug log so startup checks are visible in the main log
                try { File.AppendAllText(DebugLogFile, line); } catch { }
            }
            catch { /* don't break startup for logging failures */ }
        }

        private void HandleCarryOverIfNewDay()
        {
            try
            {
                MetaInfo meta = null;
                if (File.Exists(MetaFileName))
                {
                    var jm = File.ReadAllText(MetaFileName);
                    meta = JsonSerializer.Deserialize<MetaInfo>(jm);
                }

                var lastOpenedDate = meta?.LastOpened;
                var previousStartString = lastOpenedDate.HasValue ? lastOpenedDate.Value.ToString("yyyy-MM-dd HH:mm:ss") : "--";
                var today = DateTime.Today;

                // If meta missing, attempt to infer date from saved keys
                if (!lastOpenedDate.HasValue)
                {
                    try
                    {
                        var candidateDates = AllTasks.Keys
                            .Select(k => { DateTime dt; return new { Ok = DateTime.TryParseExact(k, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dt), Date = dt }; })
                            .Where(x => x.Ok)
                            .Select(x => x.Date)
                            .ToList();
                        if (candidateDates.Count > 0)
                        {
                            lastOpenedDate = candidateDates.Max();
                        }
                    }
                    catch { }
                }

                // Update last opened display
                try { UpdateLastOpenedText(lastOpenedDate); } catch { }

                // Save current open time for next run
                var newMeta = new MetaInfo { LastOpened = DateTime.Now };
                try { File.WriteAllText(MetaFileName, JsonSerializer.Serialize(newMeta)); } catch { }

                bool carryoverCalled = false;
                string reason = "";
                bool? dialogResult = null;

                if (!lastOpenedDate.HasValue)
                {
                    reason = "NoPreviousStartFound";
                }
                else if (lastOpenedDate.Value.Date >= today)
                {
                    reason = "PreviousStartIsSameOrLaterThanToday";
                }
                else
                {
                    var lastOpenedDateOnly = lastOpenedDate.Value.Date;
                    var key = DateKey(lastOpenedDateOnly);
                    if (!AllTasks.ContainsKey(key))
                    {
                        reason = "NoTasksForPreviousDate";
                    }
                    else
                    {
                        var incomplete = AllTasks[key].Where(t => !t.IsComplete).ToList();
                        if (incomplete.Count == 0)
                        {
                            reason = "NoIncompleteTasks";
                        }
                        else
                        {
                            // Call the dialog
                            carryoverCalled = true;
                            try
                            {
                                var dlg = new CarryOverDialog(incomplete) { Owner = this };
                                dialogResult = dlg.ShowDialog() == true;
                                reason = dialogResult == true ? "DialogShownAndProcessed" : "DialogShownButCancelled";

                                if (dialogResult == true)
                                {
                                    foreach (var item in dlg.Items)
                                    {
                                        var t = item.Task;
                                        if (item.IsMarkCompleted)
                                        {
                                            var found = AllTasks[key].FirstOrDefault(x => x.Id == t.Id);
                                            if (found != null) found.IsComplete = true;
                                        }
                                        else if (item.IsCopyFuture && item.FutureDate.HasValue)
                                        {
                                            var newKey = DateKey(item.FutureDate.Value.Date);
                                            var copy = new TaskModel(t.TaskName, false, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), true, item.FutureDate, t.LinkPath, t.Id);
                                            if (!AllTasks.ContainsKey(newKey)) AllTasks[newKey] = new List<TaskModel>();
                                            AllTasks[newKey].Add(copy);
                                        }
                                        else
                                        {
                                            var todayKey = DateKey(today);
                                            var copy = new TaskModel(t.TaskName, false, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), false, null, t.LinkPath, t.Id);
                                            if (!AllTasks.ContainsKey(todayKey)) AllTasks[todayKey] = new List<TaskModel>();
                                            AllTasks[todayKey].Add(copy);
                                        }
                                    }

                                    if (_currentDate == DateTime.Today)
                                        LoadTasksForDate(DateTime.Today);
                                }
                            }
                            catch (Exception ex)
                            {
                                reason = "DialogError:" + ex.Message.Replace('\n', ' ').Replace('\r', ' ');
                            }
                        }
                    }
                }

                // Single consolidated log entry for this startup/carryover check
                AppendStartupLog($"StartupAttempt: AttemptTime={DateTime.Now:yyyy-MM-dd HH:mm:ss}, PreviousStart={previousStartString}, CarryoverCalled={carryoverCalled}, Result={reason}{(dialogResult.HasValue ? ", DialogResult=" + dialogResult.Value.ToString() : "")} ");
            }
            catch (Exception ex)
            {
                AppendStartupLog($"StartupCheckError: {ex.Message}");
            }
        }

        private void Current_Exit(object sender, ExitEventArgs e)
        {
            CommitAllTaskEdits();
            SaveTasks();
            try { TimeTrackingService.Instance.RecordClose(); } catch { }
        }

        private void CommitAllTaskEdits()
        {
            foreach (var item in lbTasksList.Items)
            {
                var listViewItem = lbTasksList.ItemContainerGenerator.ContainerFromItem(item) as ListViewItem;
                if (listViewItem != null)
                {
                    var textBox = FindVisualChild<TextBox>(listViewItem);
                    if (textBox != null && textBox.IsFocused)
                    {
                        textBox.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                    }
                }
            }
        }

        private void LoadTasks()
        {
            try
            {
                if (File.Exists(SaveFileName))
                {
                    var json = File.ReadAllText(SaveFileName);
                    var items = JsonSerializer.Deserialize<Dictionary<string, List<TaskModel>>>(json);
                    if (items != null)
                    {
                        LogLoadedTasks(items);
                        foreach (var key in items.Keys.ToList()) items[key] = items[key].Where(t => t != null && !t.IsPlaceholder).ToList();
                        AllTasks = items;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load tasks: {ex.Message}", "Error");
            }
            UpdateTotals();
        }

        private void SaveTasks()
        {
            if (_isInitializing) return;
            LogAllTasks();
            try
            {
                if (_mode == ViewMode.Today) SaveCurrentDateTasks();

                var sanitized = new Dictionary<string, List<TaskModel>>();
                foreach (var kv in AllTasks) sanitized[kv.Key] = kv.Value.Where(t => t != null && !t.IsPlaceholder).ToList();

                try { Directory.CreateDirectory(DataDirectory); } catch { }
                var json = JsonSerializer.Serialize(sanitized, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SaveFileName, json);
                try { UpdateLastSavedText(DateTime.Now); } catch { }
                try { UpdateTotals(); } catch { }
            }
            catch (Exception ex) { MessageBox.Show($"Failed to save tasks: {ex.Message}", "Error"); }
        }

        private void LogLoadedTasks(Dictionary<string, List<TaskModel>> items)
        {
            try
            {
                try { Directory.CreateDirectory(DataDirectory); } catch { }
                var sb = new StringBuilder();
                sb.AppendLine($"[{DateTime.Now}] Loaded AllTasks:");
                foreach (var kv in items)
                {
                    sb.AppendLine($"Date: {kv.Key}");
                    foreach (var t in kv.Value) sb.AppendLine($"  - {t.TaskName} (Complete: {t.IsComplete}, Placeholder: {t.IsPlaceholder})");
                }
                File.AppendAllText(DebugLogFile, sb.ToString());
            }
            catch { }
        }

        private void LogAllTasks()
        {
            try
            {
                try { Directory.CreateDirectory(DataDirectory); } catch { }
                var sb = new StringBuilder();
                sb.AppendLine($"[{DateTime.Now}] Saving AllTasks:");
                foreach (var kv in AllTasks)
                {
                    sb.AppendLine($"Date: {kv.Key}");
                    foreach (var t in kv.Value) sb.AppendLine($"  - {t.TaskName} (Complete: {t.IsComplete}, Placeholder: {t.IsPlaceholder})");
                }
                File.AppendAllText(DebugLogFile, sb.ToString());
            }
            catch { }
        }

        private string DateKey(DateTime dt) => dt.ToString("yyyy-MM-dd");

        private void SetCurrentDate(DateTime dt)
        {
            _currentDate = dt.Date;
            CurrentDayButton.Content = FormatDate(_currentDate);
            LoadTasksForDate(_currentDate);
            bool isToday = _currentDate == DateTime.Today;
            foreach (var task in TaskList) task.IsReadOnly = !isToday;
            UpdateTitle();
            UpdateJumpTodayIcon();
            ApplyTodaySearchFilter();
        }

        private void UpdateJumpTodayIcon()
        {
            try
            {
                if (JumpTodayDot != null && JumpTodayCheck != null)
                {
                    if (_currentDate == DateTime.Today) { JumpTodayDot.Visibility = Visibility.Collapsed; JumpTodayCheck.Visibility = Visibility.Visible; }
                    else { JumpTodayDot.Visibility = Visibility.Visible; JumpTodayCheck.Visibility = Visibility.Collapsed; }
                }
            }
            catch { }
        }

        private void LoadTasksForDate(DateTime dt)
        {
            TaskList.Clear();
            var key = DateKey(dt);
            if (AllTasks.ContainsKey(key))
            {
                foreach (var t in AllTasks[key])
                {
                    var model = new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id)
                    {
                        IsReadOnly = dt != DateTime.Today,
                        InRepository = false,
                        PreferredTodayIndex = t.PreferredTodayIndex
                    };
                    TaskList.Add(model);
                }
            }
            if (dt == DateTime.Today) EnsureHasPlaceholder();
            UpdateTotals();
            ApplyTodaySearchFilter();
        }

        private void EnsureHasPlaceholder()
        {
            if (TaskList.Count == 0 || !TaskList.Last().IsPlaceholder) TaskList.Add(new TaskModel(string.Empty, false, true));
            for (int i = 0; i < TaskList.Count - 1; i++) TaskList[i].IsPlaceholder = false;
        }

        private string GetOrdinalSuffix(int day)
        {
            if (day % 100 >= 11 && day % 100 <= 13) return "th";
            switch (day % 10) { case 1: return "st"; case 2: return "nd"; case 3: return "rd"; default: return "th"; }
        }
        private string FormatDate(DateTime dt)
        {
            int day = dt.Day; string suffix = GetOrdinalSuffix(day);
            return $"{dt:dddd}, {day}{suffix} {dt:MMMM}, {dt:yyyy}";
        }
        private string FormatDateShort(DateTime dt)
        {
            int day = dt.Day; string suffix = GetOrdinalSuffix(day);
            return $"{day}{suffix} {dt:MMMM} {dt:yyyy}";
        }

        private void UpdateTitle()
        {
            switch (_mode)
            {
                case ViewMode.Today:
                    Title = _currentDate == DateTime.Today ? $"Todo | Today - {FormatDateShort(_currentDate)}" : $"Todo | {FormatDateShort(_currentDate)}"; break;
                case ViewMode.Repository:
                    Title = "Todo | Repository"; break;
                case ViewMode.People:
                    string person = PeopleFilterComboBox?.SelectedItem as string;
                    Title = string.IsNullOrEmpty(person) ? "Todo | Tasks with Other People" : $"Todo | Tasks with {person}"; break;
                case ViewMode.Meetings:
                    string meeting = MeetingsFilterComboBox?.SelectedItem as string;
                    Title = string.IsNullOrEmpty(meeting) ? "Todo | All Meetings" : $"Todo | Meeting {meeting}"; break;
                case ViewMode.All:
                    string term = SearchTextBox?.Text;
                    Title = string.IsNullOrWhiteSpace(term) ? "Todo | All Tasks" : $"Todo | Tasks containing {term}"; break;
                default:
                    Title = "Todo"; break;
            }
        }

        // Navigation
        private void PreviousDayButton_Click(object sender, RoutedEventArgs e) => SetCurrentDate(_currentDate.AddDays(-1));
        private void NextDayButton_Click(object sender, RoutedEventArgs e) => SetCurrentDate(_currentDate.AddDays(1));
        private void CalendarButton_Click(object sender, RoutedEventArgs e)
        {
            CalendarButton.Focus();
            CalendarPopup.IsOpen = true;
            CalendarControl.SelectedDate = _currentDate;
            Dispatcher.BeginInvoke(new Action(() => CalendarControl.Focus()), DispatcherPriority.Input);
        }
        private void JumpToTodayButton_Click(object sender, RoutedEventArgs e)
        {
            try { if (CalendarPopup != null) CalendarPopup.IsOpen = false; } catch { }
            SetCurrentDate(DateTime.Today);
            try { if (CalendarControl != null) CalendarControl.SelectedDate = DateTime.Today; } catch { }
        }
        private void CalendarControl_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CalendarControl.SelectedDate.HasValue)
            {
                SetCurrentDate(CalendarControl.SelectedDate.Value);
                CalendarPopup.IsOpen = false;
            }
        }

        private enum ViewMode { Today, Repository, People, Meetings, All }
        private ViewMode _mode = ViewMode.Today;
        private string _activeFilter = null;

        private void TodayButton_Click(object sender, RoutedEventArgs e)
        {
            DateControlsGrid.Visibility = Visibility.Visible;
            PeopleFilterComboBox.Visibility = Visibility.Collapsed; MeetingsFilterComboBox.Visibility = Visibility.Collapsed;
            PeopleFilterContainer.Visibility = Visibility.Collapsed; MeetingsFilterContainer.Visibility = Visibility.Collapsed;
            SearchTextBox.Visibility = Visibility.Collapsed; SearchBoxBorder.Visibility = Visibility.Collapsed;
            _mode = ViewMode.Today; this.Tag = "Today"; _activeFilter = null; FilterPopup.IsOpen = false; CalendarPopup.IsOpen = false;
            TodaySearchContainer.Visibility = Visibility.Visible;
            ClearTodaySearchFilter();
            SetCurrentDate(DateTime.Today);
            ApplyTodaySearchFilter();
        }
        private void RepositoryButton_Click(object sender, RoutedEventArgs e)
        {
            _mode = ViewMode.Repository; this.Tag = "Repository";
            DateControlsGrid.Visibility = Visibility.Collapsed;
            PeopleFilterContainer.Visibility = Visibility.Collapsed; MeetingsFilterContainer.Visibility = Visibility.Collapsed;
            SearchTextBox.Visibility = Visibility.Collapsed; SearchBoxBorder.Visibility = Visibility.Collapsed;
            TodaySearchContainer.Visibility = Visibility.Collapsed;
            ApplyRepositoryFilter();
            UpdateTitle();
        }
        private void PeopleButton_Click(object sender, RoutedEventArgs e)
        {
            _mode = ViewMode.People; this.Tag = "People";
            DateControlsGrid.Visibility = Visibility.Collapsed;
            PeopleFilterContainer.Visibility = Visibility.Visible; PeopleFilterComboBox.Visibility = Visibility.Visible;
            SearchTextBox.Visibility = Visibility.Collapsed; SearchBoxBorder.Visibility = Visibility.Collapsed;
            MeetingsFilterContainer.Visibility = Visibility.Collapsed; TodaySearchContainer.Visibility = Visibility.Collapsed;
            ClearTodaySearchFilter();
            PopulatePeopleComboBox();
            try { ApplyPeopleFilter(null); } catch { ApplyFilter(); }
            UpdateTitle();
        }
        private void MeetingButton_Click(object sender, RoutedEventArgs e)
        {
            _mode = ViewMode.Meetings; this.Tag = "Meetings";
            DateControlsGrid.Visibility = Visibility.Collapsed;
            MeetingsFilterContainer.Visibility = Visibility.Visible; MeetingsFilterComboBox.Visibility = Visibility.Visible;
            SearchTextBox.Visibility = Visibility.Collapsed; SearchBoxBorder.Visibility = Visibility.Collapsed;
            PeopleFilterContainer.Visibility = Visibility.Collapsed; TodaySearchContainer.Visibility = Visibility.Collapsed;
            ClearTodaySearchFilter();
            PopulateMeetingsComboBox();
            try { ApplyMeetingsFilter(null); } catch { ApplyFilter(); }
            UpdateTitle();
        }
        private void AllButton_Click(object sender, RoutedEventArgs e)
        {
            _mode = ViewMode.All; this.Tag = "All"; FilterPopup.IsOpen = false; CalendarPopup.IsOpen = false;
            DateControlsGrid.Visibility = Visibility.Collapsed;
            PeopleFilterContainer.Visibility = Visibility.Collapsed; MeetingsFilterContainer.Visibility = Visibility.Collapsed;
            SearchBoxBorder.Visibility = Visibility.Visible; SearchTextBox.Visibility = Visibility.Visible; SearchTextBox.Text = string.Empty; SearchTextBox.Focus();
            TodaySearchContainer.Visibility = Visibility.Collapsed; ClearTodaySearchFilter();
            try { ApplyAllFilter(SearchTextBox?.Text); } catch { ApplyFilter(); }
            UpdateTitle();
        }

        private static T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child is T t) return t;
                var childOfChild = FindVisualChild<T>(child);
                if (childOfChild != null) return childOfChild;
            }
            return null;
        }

        private void SaveCurrentDateTasks()
        {
            var key = DateKey(_currentDate);
            var realTasks = TaskList.Where(t => !t.IsPlaceholder)
                .Select(t => new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id)
                { PreferredTodayIndex = t.PreferredTodayIndex })
                .ToList();

            if (realTasks.Count > 0) AllTasks[key] = realTasks; else if (AllTasks.ContainsKey(key)) AllTasks.Remove(key);
        }

        private void TaskOptionsButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is TaskModel task && !task.IsPlaceholder)
            {
                // capture origin info
                var originKey = (_mode == ViewMode.Repository) ? RepositoryKey : DateKey(_currentDate);
                var originalId = task.Id;

                // Gather suggestions from all tasks (exclude repository if needed?)
                var allPeople = AllTasks.Where(kv => kv.Key != RepositoryKey).SelectMany(kv => kv.Value.SelectMany(t => t.People)).Distinct().ToList();
                var allMeetings = AllTasks.Where(kv => kv.Key != RepositoryKey).SelectMany(kv => kv.Value.SelectMany(t => t.Meetings)).Distinct().ToList();
                var dialog = new TaskDialog(task, allPeople, allMeetings) { Owner = this };
                if (dialog.ShowDialog() == true)
                {
                    // update task properties from dialog (UI model)
                    task.TaskName = dialog.TaskTitle;
                    task.Description = dialog.TaskDescription;
                    task.People = new List<string>(dialog.TaskPeople);
                    task.Meetings = new List<string>(dialog.TaskMeetings);

                    // ensure link is updated from dialog
                    task.LinkPath = dialog.LinkPath;

                    // handle future scheduling on UI model
                    task.IsFuture = dialog.IsFuture;
                    task.FutureDate = dialog.FutureDate;

                    var todayKey = DateKey(DateTime.Today);
                    string targetKey = null;

                    // Find where this task currently exists in storage (preserve ordering/index)
                    var storedEntries = AllTasks
                        .SelectMany(kv => kv.Value.Select((t, idx) => new { Key = kv.Key, Task = t, Index = idx }))
                        .Where(x => x.Task.Id == originalId)
                        .ToList();

                    // Determine desired target key:
                    // - If user marked as future with a date, use that date
                    // - Else, if task already exists in storage, keep its existing date (don't move to today)
                    // - Else default to today
                    if (task.IsFuture && task.FutureDate.HasValue)
                    {
                        targetKey = DateKey(task.FutureDate.Value.Date);
                    }
                    else if (storedEntries.Any())
                    {
                        // Keep the first existing date to preserve original placement unless user explicitly moved the date
                        targetKey = storedEntries.First().Key;
                    }
                    else
                    {
                        targetKey = todayKey;
                    }

                    // Update or move in storage while preserving ordering when possible
                    // 1) Update any stored entries that already live in targetKey in place
                    bool updatedInPlace = false;
                    if (AllTasks.ContainsKey(targetKey))
                    {
                        var list = AllTasks[targetKey];
                        for (int i = 0; i < list.Count; i++)
                        {
                            if (list[i].Id == originalId)
                            {
                                // Update properties in place to preserve ordering
                                list[i].TaskName = task.TaskName;
                                list[i].Description = task.Description;
                                list[i].People = new List<string>(task.People ?? new List<string>());
                                list[i].Meetings = new List<string>(task.Meetings ?? new List<string>());
                                list[i].LinkPath = task.LinkPath;
                                list[i].IsFuture = task.IsFuture;
                                list[i].FutureDate = task.FutureDate;
                                list[i].IsComplete = task.IsComplete;
                                list[i].PreferredTodayIndex = task.PreferredTodayIndex;
                                updatedInPlace = true;
                                // If the UI model has SetDate/ShowDate, leave them to be derived when reloading
                            }
                        }
                    }

                    // 2) Remove this task from any other buckets where it previously existed (if moving)
                    foreach (var k in AllTasks.Keys.ToList())
                    {
                        if (k == targetKey) continue; // keep target bucket
                        var countBefore = AllTasks[k].Count;
                        AllTasks[k].RemoveAll(t => t.Id == originalId);
                        if (AllTasks[k].Count == 0) AllTasks.Remove(k);
                    }

                    // 3) If we didn't update in place, add to the target bucket (preserve existing relative ordering not possible here)
                    if (!updatedInPlace)
                    {
                        var copy = new TaskModel(task.TaskName, task.IsComplete, false, task.Description, new List<string>(task.People ?? new List<string>()), new List<string>(task.Meetings ?? new List<string>()), task.IsFuture, task.FutureDate, task.LinkPath, task.Id)
                        {
                            PreferredTodayIndex = task.PreferredTodayIndex
                        };
                        if (!AllTasks.ContainsKey(targetKey)) AllTasks[targetKey] = new List<TaskModel>();

                        AllTasks[targetKey].Add(copy);

                        // If the task was visible in the current TaskList (we were viewing that date), remove it if it moved away
                        if (originKey != targetKey)
                        {
                            TaskList.Remove(task);
                            if (_mode == ViewMode.Today && _currentDate == DateTime.ParseExact(targetKey, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture))
                            {
                                // UI model 'task' already has the updated values from dialog above.
                            }
                        }
                    }

                    // Save after editing a task (dialog closed)
                    SaveTasks();
                }
            }
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            // For now, search behaves like Today (clears filters).
            _mode = ViewMode.Today;
            this.Tag = "Today";
            FilterPopup.IsOpen = false;
            CalendarPopup.IsOpen = false;
            SetCurrentDate(DateTime.Today);
        }

        private void FilterListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FilterListBox.SelectedItem is string s)
            {
                _activeFilter = s;
                ApplyFilter();
                FilterPopup.IsOpen = false;
                UpdateTitle();
            }
        }

        private void ApplyFilter()
        {
            TaskList.Clear();

            IEnumerable<TaskModel> tasks = AllTasks.Where(kv => kv.Key != RepositoryKey).SelectMany(list => list.Value);

            // Apply active filter
            if (!string.IsNullOrEmpty(_activeFilter))
            {
                tasks = tasks.Where(t => t.Meetings != null && t.Meetings.Contains(_activeFilter));
            }

            // Deduplicate by Id
            var unique = tasks.GroupBy(t => t.Id).Select(g => g.First()).ToList();

            foreach (var t in unique)
            {
                TaskList.Add(new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id));
            }

            UpdateTotals();
        }

        // People combobox selection handler
        private void PeopleFilterComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PeopleFilterComboBox.SelectedItem is string s && string.IsNullOrEmpty(s))
            {
                ApplyPeopleFilter(null);
            }
            else if (PeopleFilterComboBox.SelectedItem is string name)
            {
                ApplyPeopleFilter(name);
            }
            UpdateTitle();
        }

        // Meetings combobox selection handler
        private void MeetingsFilterComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MeetingsFilterComboBox.SelectedItem is string s && string.IsNullOrEmpty(s))
            {
                ApplyMeetingsFilter(null);
            }
            else if (MeetingsFilterComboBox.SelectedItem is string name)
            {
                ApplyMeetingsFilter(name);
            }
            UpdateTitle();
        }

        // Search textbox change handler
        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var tb = sender as TextBox;
            ApplyAllFilter(tb?.Text);
            UpdateTitle();
        }

        // New helper: determine the date this task should be considered associated with (future date if set, otherwise the storage key date)
        private DateTime? GetAssociatedDate(TaskModel t, string dateKey)
        {
            try
            {
                if (t != null && t.IsFuture && t.FutureDate.HasValue)
                    return t.FutureDate.Value.Date;
                if (!string.IsNullOrEmpty(dateKey) && DateTime.TryParseExact(dateKey, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var dt))
                    return dt.Date;
            }
            catch { }
            return null;
        }

        // People filtering: default to only show tasks due today or in future; toggle shows older/completed tasks
        private void ApplyPeopleFilter(string person)
        {
            TaskList.Clear();
            // Include date key so we can filter by task date
            var entries = AllTasks.Where(kv => kv.Key != RepositoryKey).SelectMany(kv => kv.Value.Select(t => new { Task = t, DateKey = kv.Key }));
            var tasks = entries.Where(e => e.Task.People != null && e.Task.People.Count > 0).Select(e => new { e.Task, e.DateKey });
            if (!string.IsNullOrEmpty(person))
                tasks = tasks.Where(e => e.Task.People.Contains(person));

            bool showAll = PeopleShowAllToggle?.IsChecked == true;
            if (!showAll)
            {
                tasks = tasks.Where(e =>
                {
                    var assoc = GetAssociatedDate(e.Task, e.DateKey);
                    // include tasks that are today or in the future
                    return assoc.HasValue ? assoc.Value.Date >= DateTime.Today : true;
                });
            }

            var unique = tasks.GroupBy(e => e.Task.Id).Select(g => g.First().Task).ToList();
            foreach (var t in unique)
            {
                TaskList.Add(new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id));
            }

            UpdateTotals();
        }

        private void ApplyMeetingsFilter(string meeting)
        {
            TaskList.Clear();
            var entries = AllTasks.Where(kv => kv.Key != RepositoryKey).SelectMany(kv => kv.Value.Select(t => new { Task = t, DateKey = kv.Key }));
            var tasks = entries.Where(e => e.Task.Meetings != null && e.Task.Meetings.Count > 0).Select(e => new { e.Task, e.DateKey });
            if (!string.IsNullOrEmpty(meeting))
                tasks = tasks.Where(e => e.Task.Meetings.Contains(meeting));

            bool showAll = MeetingsShowAllToggle?.IsChecked == true;
            if (!showAll)
            {
                tasks = tasks.Where(e =>
                {
                    var assoc = GetAssociatedDate(e.Task, e.DateKey);
                    return assoc.HasValue ? assoc.Value.Date >= DateTime.Today : true;
                });
            }

            var unique = tasks.GroupBy(e => e.Task.Id).Select(g => g.First().Task).ToList();
            foreach (var t in unique)
            {
                TaskList.Add(new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id));
            }

            UpdateTotals();
        }

        private void ApplyAllFilter(string term)
        {
            TaskList.Clear();

            // Build a sequence of tasks annotated with their source key (date or repository)
            var entries = AllTasks
                .SelectMany(kv => kv.Value.Select(t => new { Task = t, DateKey = kv.Key }))
                .ToList();

            if (!string.IsNullOrEmpty(term))
            {
                var lower = term.ToLowerInvariant();
                entries = entries.Where(e => (!string.IsNullOrEmpty(e.Task.TaskName) && e.Task.TaskName.ToLowerInvariant().Contains(lower)) || (!string.IsNullOrEmpty(e.Task.Description) && e.Task.Description.ToLowerInvariant().Contains(lower))).ToList();
            }

            // Deduplicate by Id, keeping the first occurrence
            var unique = entries.GroupBy(e => e.Task.Id).Select(g => g.First()).ToList();
            foreach (var e in unique)
            {
                var t = e.Task;
                var model = new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id);
                if (e.DateKey == RepositoryKey)
                {
                    // Show repository label instead of date
                    model.InRepository = true;
                    model.SetDate = null;
                    model.ShowDate = true;
                }
                else
                {
                    // Set the date this task was stored under
                    if (DateTime.TryParseExact(e.DateKey, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var dt))
                    {
                        model.SetDate = dt.Date;
                    }
                    else
                    {
                        model.SetDate = null;
                    }
                    // In All mode we want to show the date under each item
                    model.ShowDate = true;
                }

                TaskList.Add(model);
            }

            UpdateTotals();
        }

        private void ApplyRepositoryFilter()
        {
            TaskList.Clear();
            if (AllTasks.ContainsKey(RepositoryKey))
            {
                foreach (var t in AllTasks[RepositoryKey])
                {
                    var model = new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id)
                    { InRepository = true, PreferredTodayIndex = t.PreferredTodayIndex };
                    TaskList.Add(model);
                }
            }
            UpdateTotals();
        }

        private void LbTasksList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (lbTasksList.SelectedItem is TaskModel tm)
                {
                    if (!string.IsNullOrWhiteSpace(tm.LinkPath))
                    {
                        var path = tm.LinkPath;
                        if (Directory.Exists(path))
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true, Verb = "open" });
                        }
                        else if (File.Exists(path))
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true, Verb = "open" });
                        }
                    }
                }
            }
            catch { }
        }

        private void OpenLinkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button btn && btn.DataContext is TaskModel tm)
                {
                    var path = tm?.LinkPath; if (string.IsNullOrWhiteSpace(path)) return;
                    try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true }); }
                    catch
                    {
                        try
                        {
                            if (Uri.TryCreate(path, UriKind.Absolute, out var u)) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = u.ToString(), UseShellExecute = true });
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        private void BackupNowButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                var backupName = $"{timestamp}_tasks_backup.json";
                try { Directory.CreateDirectory(DataDirectory); } catch { }
                var backupPath = Path.Combine(DataDirectory, backupName);

                if (File.Exists(SaveFileName)) File.Copy(SaveFileName, backupPath, overwrite: true);
                else
                {
                    var sanitized = new Dictionary<string, List<TaskModel>>();
                    foreach (var kv in AllTasks) sanitized[kv.Key] = kv.Value.Where(t => t != null && !t.IsPlaceholder).ToList();
                    var json = JsonSerializer.Serialize(sanitized, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(backupPath, json);
                }
                MessageBox.Show($"Backup saved: {backupPath}", "Backup", MessageBoxButton.OK, MessageBoxImage.Information);
                UpdateLastBackupText(DateTime.Now);
                UpdateTotals();
            }
            catch (Exception ex) { MessageBox.Show($"Backup failed: {ex.Message}", "Backup", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private void AutoBackupNow()
        {
            try
            {
                var folder = AutoBackupFolder;
                try { Directory.CreateDirectory(folder); } catch { }
                var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                var backupName = $"{timestamp}_tasks_backup.json";
                var backupPath = Path.Combine(folder, backupName);

                if (File.Exists(SaveFileName)) File.Copy(SaveFileName, backupPath, overwrite: true);
                else
                {
                    var sanitized = new Dictionary<string, List<TaskModel>>();
                    foreach (var kv in AllTasks) sanitized[kv.Key] = kv.Value.Where(t => t != null && !t.IsPlaceholder).ToList();
                    var json = JsonSerializer.Serialize(sanitized, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(backupPath, json);
                }

                try { UpdateLastBackupText(DateTime.Now); } catch { }

                try
                {
                    var files = Directory.GetFiles(folder, "*_tasks_backup.json").OrderByDescending(f => Path.GetFileName(f)).ToList();
                    foreach (var f in files.Skip(AutoBackupKeep)) { try { File.Delete(f); } catch { } }
                }
                catch (Exception ex) { try { File.AppendAllText(DebugLogFile, $"[{DateTime.Now}] AutoBackup trim error: {ex.Message}\n"); } catch { } }
            }
            catch (Exception ex) { try { File.AppendAllText(DebugLogFile, $"[{DateTime.Now}] AutoBackup error: {ex.Message}\n"); } catch { } }
        }

        private DateTime? FindLatestBackupTimestamp()
        {
            try
            {
                var candidates = new List<DateTime>();
                if (Directory.Exists(AutoBackupFolder))
                {
                    foreach (var f in Directory.GetFiles(AutoBackupFolder, "*_tasks_backup.json"))
                    {
                        var name = Path.GetFileNameWithoutExtension(f);
                        var parts = name.Split(new[] { "_tasks_backup" }, StringSplitOptions.None);
                        if (parts.Length > 0 && DateTime.TryParseExact(parts[0], "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var dt)) candidates.Add(dt);
                    }
                }
                if (Directory.Exists(DataDirectory))
                {
                    foreach (var f in Directory.GetFiles(DataDirectory, "*_tasks_backup.json"))
                    {
                        var name = Path.GetFileNameWithoutExtension(f);
                        var parts = name.Split(new[] { "_tasks_backup" }, StringSplitOptions.None);
                        if (parts.Length > 0 && DateTime.TryParseExact(parts[0], "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var dt)) candidates.Add(dt);
                    }
                }
                if (candidates.Count == 0) return null; return candidates.Max();
            }
            catch { return null; }
        }

        private void UpdateLastBackupText(DateTime? dt)
        {
            try { (FindName("LastBackupDataTextBlock") as TextBlock)!.Text = dt.HasValue ? dt.Value.ToString("yyyy-MM-dd HH:mm:ss") : "--"; } catch { }
        }
        private void UpdateLastSavedText(DateTime? dt)
        {
            try { (FindName("LastSavedDataTextBlock") as TextBlock)!.Text = dt.HasValue ? dt.Value.ToString("yyyy-MM-dd HH:mm:ss") : "--"; } catch { }
        }
        private void UpdateLastOpenedText(DateTime? dt)
        {
            try { (FindName("LastOpenedDataTextBlock") as TextBlock)!.Text = dt.HasValue ? dt.Value.ToString("yyyy-MM-dd HH:mm:ss") : "--"; } catch { }
        }
        private void UpdateTotals()
        {
            try
            {
                var totalTb = FindName("TotalTasksDataTextBlock") as TextBlock;
                var viewTb = FindName("CurrentViewCountDataTextBlock") as TextBlock;
                if (totalTb == null || viewTb == null) return;
                int total = 0; try { total = AllTasks.Values.SelectMany(list => list).Count(t => t != null && !t.IsPlaceholder); } catch { total = 0; }
                int current = 0; try { current = lbTasksList?.Items.OfType<TaskModel>().Count(t => t != null && !t.IsPlaceholder) ?? TaskList.Count(t => t != null && !t.IsPlaceholder); } catch { current = 0; }
                totalTb.Text = total.ToString(); viewTb.Text = current.ToString();
            }
            catch { }
        }

        private bool IsOldCompleted(TaskModel t, string dateKey)
        {
            if (t == null || !t.IsComplete || string.IsNullOrEmpty(dateKey)) return false;
            if (!DateTime.TryParseExact(dateKey, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var dt)) return false;
            // Completed on or before yesterday should be hidden in People/Meetings modes
            return dt.Date <= DateTime.Today.AddDays(-1);
        }

        private void PopulatePeopleComboBox()
        {
            try
            {
                bool showAll = PeopleShowAllToggle?.IsChecked == true;
                IEnumerable<string> people;
                // Build entries with source date so we can exclude old completed items when not showing all
                var entries = AllTasks.Where(kv => kv.Key != RepositoryKey).SelectMany(kv => kv.Value.Select(t => new { Task = t, DateKey = kv.Key }));
                if (showAll)
                {
                    // include all tasks (don't filter out old completed when 'show all' is on)
                    people = entries.SelectMany(e => e.Task.People ?? new List<string>()).Distinct();
                }
                else
                {
                    // only people who have tasks due today or in the future
                    people = entries.Where(e =>
                    {
                        var assoc = GetAssociatedDate(e.Task, e.DateKey);
                        return assoc.HasValue ? assoc.Value.Date >= DateTime.Today : true;
                    })
                    .SelectMany(e => e.Task.People)
                    .Distinct();
                }
                var list = people.OrderBy(s => s).ToList();
                list.Insert(0, ""); // blank = any
                PeopleFilterComboBox.ItemsSource = list;
                PeopleFilterComboBox.SelectedIndex = 0;
            }
            catch { }
        }

        private void PopulateMeetingsComboBox()
        {
            try
            {
                bool showAll = MeetingsShowAllToggle?.IsChecked == true;
                IEnumerable<string> meetings;
                var entries = AllTasks.Where(kv => kv.Key != RepositoryKey).SelectMany(kv => kv.Value.Select(t => new { Task = t, DateKey = kv.Key }));
                if (showAll)
                {
                    // include all tasks (don't filter out old completed when 'show all' is on)
                    meetings = entries.SelectMany(e => e.Task.Meetings ?? new List<string>()).Distinct();
                }
                else
                {
                    // only meetings that have tasks due today or in the future
                    meetings = entries.Where(e =>
                    {
                        var assoc = GetAssociatedDate(e.Task, e.DateKey);
                        return assoc.HasValue ? assoc.Value.Date >= DateTime.Today : true;
                    })
                    .SelectMany(e => e.Task.Meetings)
                    .Distinct();
                }
                var list = meetings.OrderBy(s => s).ToList();
                list.Insert(0, "");
                MeetingsFilterComboBox.ItemsSource = list;
                MeetingsFilterComboBox.SelectedIndex = 0;
            }
            catch { }
        }

        private void PeopleShowAllToggle_Checked(object sender, RoutedEventArgs e) { PopulatePeopleComboBox(); if (PeopleFilterComboBox != null) ApplyPeopleFilter(PeopleFilterComboBox.SelectedItem as string); }
        private void PeopleShowAllToggle_Unchecked(object sender, RoutedEventArgs e) { PopulatePeopleComboBox(); if (PeopleFilterComboBox != null) ApplyPeopleFilter(PeopleFilterComboBox.SelectedItem as string); }
        private void MeetingsShowAllToggle_Checked(object sender, RoutedEventArgs e) { PopulateMeetingsComboBox(); if (MeetingsFilterComboBox != null) ApplyMeetingsFilter(MeetingsFilterComboBox.SelectedItem as string); }
        private void MeetingsShowAllToggle_Unchecked(object sender, RoutedEventArgs e) { PopulateMeetingsComboBox(); if (MeetingsFilterComboBox != null) ApplyMeetingsFilter(MeetingsFilterComboBox.SelectedItem as string); }

        // Today search
        private void TodaySearchTextBox_TextChanged(object sender, TextChangedEventArgs e) { ApplyTodaySearchFilter(); UpdateTotals(); }
        private void TodayClearButton_Click(object sender, RoutedEventArgs e)
        {
            try { if (TodaySearchTextBox != null) { TodaySearchTextBox.Text = string.Empty; TodaySearchTextBox.Focus(); } } catch { }
        }
        private void ApplyTodaySearchFilter()
        {
            try
            {
                var view = CollectionViewSource.GetDefaultView(lbTasksList?.ItemsSource); if (view == null) return;
                if (_mode != ViewMode.Today)
                {
                    if (view.Filter != null) { view.Filter = null; view.Refresh(); }
                    return;
                }
                var term = TodaySearchTextBox != null ? (TodaySearchTextBox.Text ?? string.Empty).Trim() : string.Empty;
                if (string.IsNullOrEmpty(term)) { if (view.Filter != null) { view.Filter = null; view.Refresh(); } return; }
                var lower = term.ToLowerInvariant();
                Predicate<object> pred = o =>
                {
                    if (o is TaskModel t)
                    {
                        if (t.IsPlaceholder) return true;
                        if (!string.IsNullOrEmpty(t.TaskName) && t.TaskName.IndexOf(lower, StringComparison.OrdinalIgnoreCase) >= 0) return true;
                        if (!string.IsNullOrEmpty(t.Description) && t.Description.IndexOf(lower, StringComparison.OrdinalIgnoreCase) >= 0) return true;
                        if (t.People != null && t.People.Any(p => !string.IsNullOrEmpty(p) && p.IndexOf(lower, StringComparison.OrdinalIgnoreCase) >= 0)) return true;
                        if (t.Meetings != null && t.Meetings.Any(m => !string.IsNullOrEmpty(m) && m.IndexOf(lower, StringComparison.OrdinalIgnoreCase) >= 0)) return true;
                        return false;
                    }
                    return true;
                };
                view.Filter = o => pred(o); view.Refresh();
            }
            catch { }
        }
        private void ClearTodaySearchFilter()
        {
            try
            {
                if (TodaySearchTextBox != null) TodaySearchTextBox.Text = string.Empty;
                var view = CollectionViewSource.GetDefaultView(lbTasksList?.ItemsSource); if (view != null) { view.Filter = null; view.Refresh(); }
            }
            catch { }
        }
        private void AllSearchClearButton_Click(object sender, RoutedEventArgs e)
        {
            try { if (SearchTextBox != null) { SearchTextBox.Text = string.Empty; SearchTextBox.Focus(); } } catch { }
        }

        // Task list interactions
        private void TaskTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (_currentDate != DateTime.Today && sender is TextBox tb) { tb.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next)); return; }
            if (sender is TextBox tb2 && tb2.DataContext is TaskModel tm && tm.IsPlaceholder) _placeholderJustFocused = true;
        }
        private void TaskTextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (sender is TextBlock tb && tb.DataContext is TaskModel tm)
                {
                    var listViewItem = lbTasksList.ItemContainerGenerator.ContainerFromItem(tm) as ListViewItem;
                    if (listViewItem != null)
                    {
                        var textBox = FindVisualChild<TextBox>(listViewItem);
                        if (textBox != null) { textBox.Focus(); textBox.CaretIndex = textBox.Text?.Length ?? 0; }
                    }
                }
            }
            catch { }
        }
        private void TaskTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (_currentDate != DateTime.Today) { e.Handled = true; return; }
            if (sender is TextBox tb && tb.DataContext is TaskModel tm && tm.IsPlaceholder)
            {
                _suppressTextChanged = true; tm.IsPlaceholder = false; tm.TaskName = ""; tb.Text = ""; _suppressTextChanged = false;
            }
        }
        private void TaskTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (sender is TextBox tb && tb.DataContext is TaskModel tm && tm.IsPlaceholder)
                {
                    bool isPasteKey = (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.V) || (Keyboard.Modifiers == ModifierKeys.Shift && e.Key == Key.Insert);
                    if (isPasteKey)
                    {
                        _suppressTextChanged = true; tm.IsPlaceholder = false; tm.TaskName = string.Empty; tb.Text = string.Empty; _suppressTextChanged = false;
                    }
                }
            }
            catch { }
        }
        private void TaskTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            try { if (sender is TextBox tb) DataObject.AddPastingHandler(tb, new DataObjectPastingEventHandler(TaskTextBox_OnPasting)); } catch { }
        }
        private void TaskTextBox_OnPasting(object sender, DataObjectPastingEventArgs e)
        {
            try { if (sender is TextBox tb && tb.DataContext is TaskModel tm && tm.IsPlaceholder) { _suppressTextChanged = true; tm.IsPlaceholder = false; tm.TaskName = string.Empty; tb.Text = string.Empty; _suppressTextChanged = false; } } catch { }
        }
        private void TaskTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressTextChanged) return;
            if (sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                if (_currentDate != DateTime.Today) { tb.Text = tm.TaskName; return; }
                if (!tm.IsPlaceholder && TaskList.Last() == tm && !string.IsNullOrWhiteSpace(tb.Text)) { _suppressTextChanged = true; tm.TaskName = tb.Text; EnsureHasPlaceholder(); _suppressTextChanged = false; }
                else if (!tm.IsPlaceholder && string.IsNullOrWhiteSpace(tb.Text) && TaskList.IndexOf(tm) != TaskList.Count - 1) { TaskList.Remove(tm); }

                if (_mode == ViewMode.Today)
                {
                    var key = DateKey(_currentDate);
                    AllTasks[key] = TaskList.Where(t => !t.IsPlaceholder)
                        .Select(t => new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id)
                        { PreferredTodayIndex = t.PreferredTodayIndex })
                        .ToList();
                }
            }
            _placeholderJustFocused = false;
            UpdateTitle();
            UpdateTotals();
        }
        private void TaskTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (_currentDate != DateTime.Today) { e.Handled = true; return; }
            if (e.Key == Key.Enter && sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                if (TaskList.Last() == tm && !tm.IsPlaceholder && !string.IsNullOrWhiteSpace(tb.Text)) EnsureHasPlaceholder();
                var idx = TaskList.IndexOf(tm); if (idx >= 0)
                {
                    var nextIdx = Math.Min(idx + 1, TaskList.Count - 1); var next = TaskList[nextIdx];
                    var listViewItem = lbTasksList.ItemContainerGenerator.ContainerFromItem(next) as ListViewItem;
                    if (listViewItem != null)
                    {
                        var nextTextBox = FindVisualChild<TextBox>(listViewItem);
                        if (nextTextBox != null) { nextTextBox.Focus(); nextTextBox.CaretIndex = nextTextBox.Text?.Length ?? 0; }
                        else Keyboard.Focus(listViewItem);
                    }
                }
                e.Handled = true;
            }
        }
        private void TaskTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_currentDate != DateTime.Today) return;
            if (sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                if (string.IsNullOrWhiteSpace(tm.TaskName)) { if (!tm.IsPlaceholder) { tm.TaskName = string.Empty; tm.IsPlaceholder = true; } }
            }
            try { if (sender is TextBox tb2) { var sv = FindVisualChild<ScrollViewer>(tb2); if (sv != null) sv.ScrollToHorizontalOffset(0); } } catch { }
            SaveTasks();
        }

        private void TimeTrackingButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new TimeTrackingDialog() { Owner = this };
                dlg.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to open time tracking dialog: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new SettingsDialog() { Owner = this };
                if (dlg.ShowDialog() == true)
                {
                    // Apply any settings that can affect current view immediately
                    // For the completion behavior, no immediate reflow needed until next toggle.
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to open settings: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ReorderTodayAfterCompletionToggle(TaskModel tm)
        {
            try
            {
                if (tm == null || tm.IsPlaceholder) return; if (_mode != ViewMode.Today) return; if (_currentDate != DateTime.Today) return;

                // Obey user setting: either move to top or do not change order
                var behavior = SettingsService.Instance.Settings.MoveBehavior;
                if (behavior == MoveCompletedBehavior.DoNotMove)
                {
                    // Keep current order but still ensure placeholder exists
                    EnsureHasPlaceholder();
                    return;
                }

                // Default MoveToTop behavior
                var oldIndex = TaskList.IndexOf(tm); if (oldIndex < 0) return;
                TaskList.RemoveAt(oldIndex);
                int placeholderIndex = TaskList.Count > 0 && TaskList.Last().IsPlaceholder ? TaskList.Count - 1 : TaskList.Count;
                int completedCount = 0; for (int i = 0; i < placeholderIndex; i++) { if (TaskList[i].IsComplete) completedCount++; else break; }
                int insertIndex = completedCount; insertIndex = Math.Min(Math.Max(0, insertIndex), TaskList.Count);
                TaskList.Insert(insertIndex, tm);
                EnsureHasPlaceholder();
            }
            catch { }
        }

        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (_mode == ViewMode.Today)
            {
                if (sender is CheckBox cb && cb.DataContext is TaskModel tm) ReorderTodayAfterCompletionToggle(tm);
                SaveTasks(); return;
            }
            if (sender is CheckBox cb2 && cb2.DataContext is TaskModel tm2)
            {
                foreach (var k in AllTasks.Keys.ToList())
                {
                    var list = AllTasks[k];
                    foreach (var stored in list.Where(x => x.Id == tm2.Id)) stored.IsComplete = tm2.IsComplete;
                }
                SaveTasks();
            }
        }

        // Repo actions
        private void SendToRepositoryButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button btn && btn.DataContext is TaskModel tm && !tm.IsPlaceholder)
                {
                    var todayKey = DateKey(_currentDate);
                    if (!AllTasks.ContainsKey(RepositoryKey)) AllTasks[RepositoryKey] = new List<TaskModel>();
                    int indexInToday = TaskList.Where(t => !t.IsPlaceholder).ToList().FindIndex(t => t.Id == tm.Id);
                    if (AllTasks.ContainsKey(todayKey)) { AllTasks[todayKey].RemoveAll(t => t.Id == tm.Id); if (AllTasks[todayKey].Count == 0) AllTasks.Remove(todayKey); }
                    var copy = new TaskModel(tm.TaskName, tm.IsComplete, false, tm.Description, new List<string>(tm.People), new List<string>(tm.Meetings), tm.IsFuture, tm.FutureDate, tm.LinkPath, tm.Id)
                    { InRepository = true, PreferredTodayIndex = indexInToday >= 0 ? (int?)indexInToday : null };
                    AllTasks[RepositoryKey].Add(copy);
                    TaskList.Remove(tm);
                    EnsureHasPlaceholder();
                    SaveTasks();
                }
            }
            catch { }
        }
        private void MoveToTodayButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button btn && btn.DataContext is TaskModel tm && !tm.IsPlaceholder)
                {
                    if (!AllTasks.ContainsKey(RepositoryKey)) return;
                    TaskModel repoStored = AllTasks[RepositoryKey].FirstOrDefault(t => t.Id == tm.Id);
                    int? preferredIndex = repoStored?.PreferredTodayIndex;
                    AllTasks[RepositoryKey].RemoveAll(t => t.Id == tm.Id); if (AllTasks[RepositoryKey].Count == 0) AllTasks.Remove(RepositoryKey);
                    var todayKey = DateKey(DateTime.Today); if (!AllTasks.ContainsKey(todayKey)) AllTasks[todayKey] = new List<TaskModel>();
                    var copy = new TaskModel(tm.TaskName, tm.IsComplete, false, tm.Description, new List<string>(tm.People), new List<string>(tm.Meetings), false, null, tm.LinkPath, tm.Id)
                    { InRepository = false, PreferredTodayIndex = preferredIndex };
                    var todayList = AllTasks[todayKey];
                    if (preferredIndex.HasValue && preferredIndex.Value >= 0 && preferredIndex.Value <= todayList.Count) todayList.Insert(preferredIndex.Value, copy); else todayList.Add(copy);
                    TaskList.Remove(tm); if (_mode == ViewMode.Today && _currentDate == DateTime.Today) LoadTasksForDate(DateTime.Today); SaveTasks();
                }
            }
            catch { }
        }

        // Expose a minimal accessor for the AllTasks dictionary to other classes (read-only snapshot)
        public static Dictionary<string, List<TaskModel>> GetAllTasks()
        {
            // Return a shallow copy to avoid external mutation
            try
            {
                var mw = Application.Current?.MainWindow as MainWindow; if (mw == null) return new Dictionary<string, List<TaskModel>>();
                return mw.AllTasks.ToDictionary(kv => kv.Key, kv => kv.Value);
            }
            catch { return new Dictionary<string, List<TaskModel>>(); }
        }

        private class MetaInfo
        {
            public DateTime? LastOpened { get; set; }
        }
    }
}

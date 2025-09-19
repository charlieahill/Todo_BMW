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
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Windows; // ensure MessageBox

namespace Todo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string SaveFileName = "tasks_by_date.json";
        private const string MetaFileName = "meta.json";

        // All tasks keyed by date string (yyyy-MM-dd)
        private Dictionary<string, List<TaskModel>> AllTasks { get; set; } = new Dictionary<string, List<TaskModel>>();

        public ObservableCollection<TaskModel> TaskList { get; set; } = new ObservableCollection<TaskModel>();

        private DateTime _currentDate = DateTime.Today;

        private bool _isInitializing = false;

        // Auto-backup support
        private DispatcherTimer _autoBackupTimer;
        private const string AutoBackupFolder = "autobackups";
        private const int AutoBackupKeep = 10;
        private readonly TimeSpan AutoBackupInterval = TimeSpan.FromMinutes(15);

        public MainWindow()
        {
            InitializeComponent();

            DataContext = this;

            _isInitializing = true;

            LoadTasks();

            SetCurrentDate(_currentDate);

            // If application was last opened on a previous day, prompt to carry over unfinished tasks
            HandleCarryOverIfNewDay();

            lbTasksList.ItemsSource = TaskList;
            lbTasksList.MouseDoubleClick += LbTasksList_MouseDoubleClick;

            Application.Current.Exit += Current_Exit;

            // Start automatic backups
            try
            {
                _autoBackupTimer = new DispatcherTimer { Interval = AutoBackupInterval };
                _autoBackupTimer.Tick += (s, e) => {
                    try { AutoBackupNow(); } catch { }
                };
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

            _isInitializing = false;

            // Record open event
            try { TimeTrackingService.Instance.RecordOpen(); } catch { }
        }

        private class MetaInfo
        {
            public DateTime? LastOpened { get; set; }
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

                var lastOpenedDate = meta?.LastOpened?.Date;

                // If meta is missing, attempt to infer the last-opened date from saved task keys.
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
                    catch { /* ignore parse errors */ }
                }

                var today = DateTime.Today;

                // update meta to now (save regardless so next run sees current open)
                var newMeta = new MetaInfo { LastOpened = DateTime.Now };
                try { File.WriteAllText(MetaFileName, JsonSerializer.Serialize(newMeta)); } catch { }

                if (lastOpenedDate.HasValue && lastOpenedDate.Value < today)
                {
                    var key = DateKey(lastOpenedDate.Value);
                    if (AllTasks.ContainsKey(key))
                    {
                        var incomplete = AllTasks[key].Where(t => !t.IsComplete).ToList();
                        if (incomplete.Count > 0)
                        {
                            var dlg = new CarryOverDialog(incomplete) { Owner = this };
                            if (dlg.ShowDialog() == true)
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
                                        var copy = new TaskModel(t.TaskName, false, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), true, item.FutureDate, t.LinkPath, Guid.NewGuid());
                                        if (!AllTasks.ContainsKey(newKey)) AllTasks[newKey] = new List<TaskModel>();
                                        AllTasks[newKey].Add(copy);
                                    }
                                    else // Copy to today (default)
                                    {
                                        var todayKey = DateKey(today);
                                        var copy = new TaskModel(t.TaskName, false, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), false, null, t.LinkPath, Guid.NewGuid());
                                        if (!AllTasks.ContainsKey(todayKey)) AllTasks[todayKey] = new List<TaskModel>();
                                        AllTasks[todayKey].Add(copy);
                                    }
                                }

                                // Do not auto-save here; saving will happen on application exit or other explicit actions.
                                if (_currentDate == DateTime.Today)
                                    LoadTasksForDate(DateTime.Today);
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void Current_Exit(object sender, ExitEventArgs e)
        {
            // Force all TaskTextBox controls to lose focus and update their bindings
            CommitAllTaskEdits();
            // Save tasks; SaveTasks will decide whether to persist current-date TaskList or just AllTasks
            SaveTasks();

            // record close event
            try { TimeTrackingService.Instance.RecordClose(); } catch { }
        }

        private void CommitAllTaskEdits()
        {
            // Find all ListViewItems in lbTasksList
            foreach (var item in lbTasksList.Items)
            {
                var listViewItem = lbTasksList.ItemContainerGenerator.ContainerFromItem(item) as ListViewItem;
                if (listViewItem != null)
                {
                    var textBox = FindVisualChild<TextBox>(listViewItem);
                    if (textBox != null && textBox.IsFocused)
                    {
                        // Move focus away to commit binding
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
                        LogLoadedTasks(items); // Log loaded tasks for debugging
                        // Remove placeholders from loaded data
                        foreach (var key in items.Keys.ToList())
                        {
                            items[key] = items[key].Where(t => t != null && !t.IsPlaceholder).ToList();
                        }
                        AllTasks = items;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load tasks: {ex.Message}", "Error");
            }
        }

        private void LogLoadedTasks(Dictionary<string, List<TaskModel>> items)
        {
            try
            {
                var logPath = "tasks_debug_log.txt";
                var sb = new StringBuilder();
                sb.AppendLine($"[{DateTime.Now}] Loaded AllTasks:");
                foreach (var kv in items)
                {
                    sb.AppendLine($"Date: {kv.Key}");
                    foreach (var t in kv.Value)
                    {
                        sb.AppendLine($"  - {t.TaskName} (Complete: {t.IsComplete}, Placeholder: {t.IsPlaceholder})");
                    }
                }
                File.AppendAllText(logPath, sb.ToString());
            }
            catch { }
        }

        private void SaveTasks()
        {
            if (_isInitializing) return; // don't save during startup

            LogAllTasks(); // Debug log before saving
            try
            {
                // Only persist the current TaskList into AllTasks when we're in Today view.
                // When in filtered views (People/Meetings/All) TaskList contains a subset of items across dates
                // and we must not overwrite the stored tasks for any date with that subset.
                if (_mode == ViewMode.Today)
                {
                    SaveCurrentDateTasks(); // Ensure current day's tasks are saved before serializing
                }

                // ensure we don't save placeholder items
                var sanitized = new Dictionary<string, List<TaskModel>>();
                foreach (var kv in AllTasks)
                {
                    sanitized[kv.Key] = kv.Value.Where(t => t != null && !t.IsPlaceholder).ToList();
                }

                var json = JsonSerializer.Serialize(sanitized, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SaveFileName, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save tasks: {ex.Message}", "Error");
            }
        }

        private void LogAllTasks()
        {
            try
            {
                var logPath = "tasks_debug_log.txt";
                var sb = new StringBuilder();
                sb.AppendLine($"[{DateTime.Now}] Saving AllTasks:");
                foreach (var kv in AllTasks)
                {
                    sb.AppendLine($"Date: {kv.Key}");
                    foreach (var t in kv.Value)
                    {
                        sb.AppendLine($"  - {t.TaskName} (Complete: {t.IsComplete}, Placeholder: {t.IsPlaceholder})");
                    }
                }
                File.AppendAllText(logPath, sb.ToString());
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
            foreach (var task in TaskList)
            {
                task.IsReadOnly = !isToday;
            }
            UpdateTitle();
            UpdateJumpTodayIcon();
            // SaveCurrentDateTasks(); // Removed to prevent wiping out tasks on startup
        }

        private void UpdateJumpTodayIcon()
        {
            try
            {
                if (JumpTodayDot != null && JumpTodayCheck != null)
                {
                    if (_currentDate == DateTime.Today)
                    {
                        JumpTodayDot.Visibility = Visibility.Collapsed;
                        JumpTodayCheck.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        JumpTodayDot.Visibility = Visibility.Visible;
                        JumpTodayCheck.Visibility = Visibility.Collapsed;
                    }
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
                    // Deep copy all properties, including id and future scheduling
                    var model = new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id);
                    model.IsReadOnly = dt != DateTime.Today;
                    TaskList.Add(model);
                }
            }
            if (dt == DateTime.Today)
            {
                EnsureHasPlaceholder();
            }
        }

        private void EnsureHasPlaceholder()
        {
            if (TaskList.Count == 0 || !TaskList.Last().IsPlaceholder)
            {
                TaskList.Add(new TaskModel(string.Empty, false, true));
            }

            // ensure only last is placeholder
            for (int i = 0; i < TaskList.Count - 1; i++)
                TaskList[i].IsPlaceholder = false;
        }

        private void SetupSampleTasks()
        {
            // not used any more
        }

        private bool _suppressTextChanged = false;
        private bool _placeholderJustFocused = false;

        // In TaskTextBox_GotFocus, prevent focus for non-today
        private void TaskTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (_currentDate != DateTime.Today && sender is TextBox tb)
            {
                tb.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                return;
            }
            if (sender is TextBox tb2 && tb2.DataContext is TaskModel tm)
            {
                if (tm.IsPlaceholder)
                {
                    _placeholderJustFocused = true;
                }
            }
        }

        private void TaskTextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // When the preview TextBlock is clicked, focus the underlying TextBox for editing
            try
            {
                if (sender is TextBlock tb && tb.DataContext is TaskModel tm)
                {
                    var listViewItem = lbTasksList.ItemContainerGenerator.ContainerFromItem(tm) as ListViewItem;
                    if (listViewItem != null)
                    {
                        var textBox = FindVisualChild<TextBox>(listViewItem);
                        if (textBox != null)
                        {
                            textBox.Focus();
                            textBox.CaretIndex = textBox.Text?.Length ?? 0;
                        }
                    }
                }
            }
            catch { }
        }

        private void TaskTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (_currentDate != DateTime.Today)
            {
                e.Handled = true;
                return;
            }
            if (sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                if (tm.IsPlaceholder)
                {
                    _suppressTextChanged = true;
                    tm.IsPlaceholder = false;
                    tm.TaskName = "";
                    tb.Text = "";
                    _suppressTextChanged = false;
                }
            }
        }

        // Handle paste via keyboard or context menu by clearing placeholder before the paste occurs
        private void TaskTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (sender is TextBox tb && tb.DataContext is TaskModel tm && tm.IsPlaceholder)
                {
                    bool isPasteKey = (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.V) || (Keyboard.Modifiers == ModifierKeys.Shift && e.Key == Key.Insert);
                    if (isPasteKey)
                    {
                        _suppressTextChanged = true;
                        tm.IsPlaceholder = false;
                        tm.TaskName = string.Empty;
                        tb.Text = string.Empty;
                        _suppressTextChanged = false;
                        // allow paste to proceed
                    }
                }
            }
            catch { }
        }

        private void TaskTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is TextBox tb)
                {
                    DataObject.AddPastingHandler(tb, new DataObjectPastingEventHandler(TaskTextBox_OnPasting));
                }
            }
            catch { }
        }

        private void TaskTextBox_OnPasting(object sender, DataObjectPastingEventArgs e)
        {
            try
            {
                if (sender is TextBox tb && tb.DataContext is TaskModel tm && tm.IsPlaceholder)
                {
                    _suppressTextChanged = true;
                    tm.IsPlaceholder = false;
                    tm.TaskName = string.Empty;
                    tb.Text = string.Empty;
                    _suppressTextChanged = false;
                }
            }
            catch { }
        }

        private void TaskTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressTextChanged) return;
            if (sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                if (_currentDate != DateTime.Today)
                {
                    tb.Text = tm.TaskName; // revert any change
                    return;
                }
                // Always add a new placeholder as soon as the last line's text changes from empty to non-empty
                if (!tm.IsPlaceholder && TaskList.Last() == tm && !string.IsNullOrWhiteSpace(tb.Text))
                {
                    _suppressTextChanged = true;
                    tm.TaskName = tb.Text;
                    EnsureHasPlaceholder();
                    _suppressTextChanged = false;
                }
                // Remove empty non-placeholder tasks (except the last)
                else if (!tm.IsPlaceholder && string.IsNullOrWhiteSpace(tb.Text) && TaskList.IndexOf(tm) != TaskList.Count - 1)
                {
                    TaskList.Remove(tm);
                }

                // update current date storage in memory only (don't persist to disk here)
                var key = DateKey(_currentDate);
                AllTasks[key] = TaskList.Where(t => !t.IsPlaceholder)
                    .Select(t => new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id))
                    .ToList();
                // Removed immediate SaveTasks();
            }
            _placeholderJustFocused = false;
            UpdateTitle();
        }

        private void TaskTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (_currentDate != DateTime.Today)
            {
                e.Handled = true;
                return;
            }

            // If Enter pressed, move focus to the next task line (create placeholder if needed)
            if (e.Key == Key.Enter && sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                // If we're on the last real item and it has text, ensure a placeholder exists before moving
                if (TaskList.Last() == tm && !tm.IsPlaceholder && !string.IsNullOrWhiteSpace(tb.Text))
                {
                    EnsureHasPlaceholder();
                }

                var idx = TaskList.IndexOf(tm);
                if (idx >= 0)
                {
                    var nextIdx = Math.Min(idx + 1, TaskList.Count - 1);
                    var next = TaskList[nextIdx];
                    var listViewItem = lbTasksList.ItemContainerGenerator.ContainerFromItem(next) as ListViewItem;
                    if (listViewItem != null)
                    {
                        var nextTextBox = FindVisualChild<TextBox>(listViewItem);
                        if (nextTextBox != null)
                        {
                            nextTextBox.Focus();
                            nextTextBox.CaretIndex = nextTextBox.Text?.Length ?? 0;
                        }
                        else
                        {
                            // Fallback: move focus programmatically
                            Keyboard.Focus(listViewItem);
                        }
                    }
                }

                e.Handled = true;
            }

            // Removed immediate SaveTasks();
        }

        private void TaskTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_currentDate != DateTime.Today)
                return;
            if (sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                if (string.IsNullOrWhiteSpace(tm.TaskName))
                {
                    if (!tm.IsPlaceholder)
                    {
                        tm.TaskName = string.Empty;
                        tm.IsPlaceholder = true;
                    }
                }
            }
            // Reset horizontal scroll so the start of the line is visible when losing focus
            try
            {
                if (sender is TextBox tb2)
                {
                    var sv = FindVisualChild<ScrollViewer>(tb2);
                    if (sv != null)
                    {
                        sv.ScrollToHorizontalOffset(0);
                      }
                  }
            }
            catch { }

            // Save after lost focus for today (persist changes)
            SaveTasks();
        }

        // Prevent checking/unchecking for non-today
        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            // If we're in Today view, follow the original behavior and persist via SaveTasks
            if (_mode == ViewMode.Today)
            {
                // Save after completion state changes for today
                SaveTasks();
                return;
            }

            // For filtered views (People/Meetings/All) the TaskList contains a subset of tasks across dates.
            // Update the corresponding task(s) in AllTasks by Id so we don't overwrite other items.
            if (sender is CheckBox cb && cb.DataContext is TaskModel tm)
            {
                // Apply the IsComplete value to all matching tasks in storage
                foreach (var k in AllTasks.Keys.ToList())
                {
                    var list = AllTasks[k];
                    foreach (var stored in list.Where(x => x.Id == tm.Id))
                    {
                        stored.IsComplete = tm.IsComplete;
                    }
                }

                // Persist the updated AllTasks
                SaveTasks();
            }
        }

        private string GetOrdinalSuffix(int day)
        {
            if (day % 100 >= 11 && day % 100 <= 13)
                return "th";
            switch (day % 10)
            {
                case 1: return "st";
                case 2: return "nd";
                case 3: return "rd";
                default: return "th";
            }
        }

        private string FormatDate(DateTime dt)
        {
            int day = dt.Day;
            string suffix = GetOrdinalSuffix(day);
            return $"{dt:dddd}, {day}{suffix} {dt:MMMM}, {dt:yyyy}";
        }

        private string FormatDateShort(DateTime dt)
        {
            int day = dt.Day;
            string suffix = GetOrdinalSuffix(day);
            return $"{day}{suffix} {dt:MMMM} {dt:yyyy}";
        }

        private void UpdateTitle()
        {
            switch (_mode)
            {
                case ViewMode.Today:
                    if (_currentDate == DateTime.Today)
                    {
                        this.Title = $"Todo | Today - {FormatDateShort(_currentDate)}";
                    }
                    else
                    {
                        this.Title = $"Todo | {FormatDateShort(_currentDate)}";
                    }
                    break;
                case ViewMode.People:
                    string person = null;
                    if (PeopleFilterComboBox != null && PeopleFilterComboBox.SelectedItem is string p)
                        person = p;
                    if (string.IsNullOrEmpty(person))
                        this.Title = "Todo | Tasks with Other People";
                    else
                        this.Title = $"Todo | Tasks with {person}";
                    break;
                case ViewMode.Meetings:
                    string meeting = null;
                    if (MeetingsFilterComboBox != null && MeetingsFilterComboBox.SelectedItem is string m)
                        meeting = m;
                    if (string.IsNullOrEmpty(meeting))
                        this.Title = "Todo | All Meetings";
                    else
                        this.Title = $"Todo | Meeting {meeting}";
                    break;
                case ViewMode.All:
                    string term = null;
                    if (SearchTextBox != null)
                        term = SearchTextBox.Text;
                    if (string.IsNullOrWhiteSpace(term))
                        this.Title = "Todo | All Tasks";
                    else
                        this.Title = $"Todo | Tasks containing {term}";
                    break;
                default:
                    this.Title = "Todo";
                    break;
            }
        }

        // Navigation handlers
        private void PreviousDayButton_Click(object sender, RoutedEventArgs e)
        {
            SetCurrentDate(_currentDate.AddDays(-1));
        }

        private void NextDayButton_Click(object sender, RoutedEventArgs e)
        {
            SetCurrentDate(_currentDate.AddDays(1));
        }

        private void CalendarButton_Click(object sender, RoutedEventArgs e)
        {
            // Workaround: set focus to button before opening popup
            CalendarButton.Focus();
            CalendarPopup.IsOpen = true;
            CalendarControl.SelectedDate = _currentDate;
            Dispatcher.BeginInvoke(new Action(() => CalendarControl.Focus()), System.Windows.Threading.DispatcherPriority.Input);
        }

        private void JumpToTodayButton_Click(object sender, RoutedEventArgs e)
        {
            // Close the calendar popup (if open) and jump to today's date
            try
            {
                if (CalendarPopup != null)
                {
                    CalendarPopup.IsOpen = false;
                }
            }
            catch { }

            SetCurrentDate(DateTime.Today);

            try
            {
                if (CalendarControl != null)
                    CalendarControl.SelectedDate = DateTime.Today;
            }
            catch { }
        }

        private void TomorrowButton_Click(object sender, RoutedEventArgs e)
        {
            var tomorrow = DateTime.Today.AddDays(1);
            // Close calendar popup if open
            try { if (CalendarPopup != null) CalendarPopup.IsOpen = false; } catch { }
            SetCurrentDate(tomorrow);
            try { if (CalendarControl != null) CalendarControl.SelectedDate = tomorrow; } catch { }
        }

        private void CalendarControl_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CalendarControl.SelectedDate.HasValue)
            {
                SetCurrentDate(CalendarControl.SelectedDate.Value);
                CalendarPopup.IsOpen = false;
            }
        }

        private enum ViewMode { Today, People, Meetings, All }
        private ViewMode _mode = ViewMode.Today;
        private string _activeFilter = null;

        private void ExitAllMode()
        {
            DateControlsGrid.Visibility = Visibility.Visible;
            if (SearchTextBox != null) SearchTextBox.Visibility = Visibility.Collapsed;
            if (SearchBoxBorder != null) SearchBoxBorder.Visibility = Visibility.Collapsed;
            _mode = ViewMode.Today;
            // hide people/meetings filter containers when exiting special modes
            if (PeopleFilterContainer != null) PeopleFilterContainer.Visibility = Visibility.Collapsed;
            if (MeetingsFilterContainer != null) MeetingsFilterContainer.Visibility = Visibility.Collapsed;
            SetCurrentDate(_currentDate);
        }

        private void TodayButton_Click(object sender, RoutedEventArgs e)
        {
            // Exit any special filter mode and return to today's view
            // restore date controls
            DateControlsGrid.Visibility = Visibility.Visible;
            // hide people/meetings filter comboboxes and containers
            if (PeopleFilterComboBox != null) PeopleFilterComboBox.Visibility = Visibility.Collapsed;
            if (MeetingsFilterComboBox != null) MeetingsFilterComboBox.Visibility = Visibility.Collapsed;
            if (PeopleFilterContainer != null) PeopleFilterContainer.Visibility = Visibility.Collapsed;
            if (MeetingsFilterContainer != null) MeetingsFilterContainer.Visibility = Visibility.Collapsed;
            // hide search UI
            if (SearchTextBox != null) SearchTextBox.Visibility = Visibility.Collapsed;
            if (SearchBoxBorder != null) SearchBoxBorder.Visibility = Visibility.Collapsed;

            _mode = ViewMode.Today;
            _activeFilter = null;
            FilterPopup.IsOpen = false;
            CalendarPopup.IsOpen = false;
            SetCurrentDate(DateTime.Today);
        }

        private void PeopleButton_Click(object sender, RoutedEventArgs e)
        {
            _mode = ViewMode.People;
            // hide date controls and show people combobox container
            DateControlsGrid.Visibility = Visibility.Collapsed;
            if (PeopleFilterContainer != null) PeopleFilterContainer.Visibility = Visibility.Visible;
            // ensure inner controls are visible (child ComboBox may have been collapsed earlier)
            if (PeopleFilterComboBox != null) PeopleFilterComboBox.Visibility = Visibility.Visible;
            if (PeopleShowAllToggle != null) { /* keep previous state; do not modify */ }
             // hide search box if visible
             if (SearchTextBox != null) SearchTextBox.Visibility = Visibility.Collapsed;
             if (SearchBoxBorder != null) SearchBoxBorder.Visibility = Visibility.Collapsed;
             // hide the meetings container when in people mode
             if (MeetingsFilterContainer != null) MeetingsFilterContainer.Visibility = Visibility.Collapsed;
             // populate people combobox according to toggle (default: only active people)
             PopulatePeopleComboBox();
             // show tasks
             try { ApplyPeopleFilter(null); } catch { ApplyFilter(); }
             UpdateTitle();
         }

         private void MeetingButton_Click(object sender, RoutedEventArgs e)
         {
             _mode = ViewMode.Meetings;
             // hide date controls and show meetings combobox
             DateControlsGrid.Visibility = Visibility.Collapsed;
             if (MeetingsFilterContainer != null) MeetingsFilterContainer.Visibility = Visibility.Visible;
             // ensure inner controls are visible
             if (MeetingsFilterComboBox != null) MeetingsFilterComboBox.Visibility = Visibility.Visible;
             if (MeetingsShowAllToggle != null) { /* keep previous state; do not modify */ }
             // hide search box if visible
             if (SearchTextBox != null) SearchTextBox.Visibility = Visibility.Collapsed;
             if (SearchBoxBorder != null) SearchBoxBorder.Visibility = Visibility.Collapsed;
             // hide the people container when in meetings mode
             if (PeopleFilterContainer != null) PeopleFilterContainer.Visibility = Visibility.Collapsed;
             // populate meetings combobox according to toggle (default: only active meetings)
             PopulateMeetingsComboBox();
             try { ApplyMeetingsFilter(null); } catch { ApplyFilter(); }
             UpdateTitle();
         }

        private void AllButton_Click(object sender, RoutedEventArgs e)
        {
            _mode = ViewMode.All;
            FilterPopup.IsOpen = false;
            CalendarPopup.IsOpen = false;

            // Show search box and hide date/filters
            DateControlsGrid.Visibility = Visibility.Collapsed;
            if (PeopleFilterContainer != null) PeopleFilterContainer.Visibility = Visibility.Collapsed;
            if (MeetingsFilterContainer != null) MeetingsFilterContainer.Visibility = Visibility.Collapsed;
            if (SearchBoxBorder != null) SearchBoxBorder.Visibility = Visibility.Visible;
            if (SearchTextBox != null)
            {
                SearchTextBox.Visibility = Visibility.Visible;
                SearchTextBox.Text = string.Empty;
                SearchTextBox.Focus();
            }

            // Load all tasks
            try { ApplyAllFilter(SearchTextBox?.Text); } catch { ApplyFilter(); }
            UpdateTitle();
        }

        // Helper to find child of type T in visual tree
        private static T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is T t)
                    return t;
                else
                {
                    T childOfChild = FindVisualChild<T>(child);
                    if (childOfChild != null)
                        return childOfChild;
                }
            }
            return null;
        }

        private void SaveCurrentDateTasks()
        {
            if (_currentDate == null) return;
            var key = DateKey(_currentDate);
            var realTasks = TaskList.Where(t => !t.IsPlaceholder)
                .Select(t => new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id))
                .ToList();

            if (realTasks.Count > 0)
            {
                AllTasks[key] = realTasks;
            }
            else
            {
                if (AllTasks.ContainsKey(key))
                    AllTasks.Remove(key);
            }
        }

        private void TaskOptionsButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is TaskModel task && !task.IsPlaceholder)
            {
                // capture origin info
                var originKey = DateKey(_currentDate);
                var originalId = task.Id;

                // Gather suggestions from all tasks
                var allPeople = AllTasks.Values.SelectMany(list => list.SelectMany(t => t.People)).Distinct().ToList();
                var allMeetings = AllTasks.Values.SelectMany(list => list.SelectMany(t => t.Meetings)).Distinct().ToList();
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
                                // ensure link path is persisted
                                list[i].LinkPath = task.LinkPath;
                                list[i].IsFuture = task.IsFuture;
                                list[i].FutureDate = task.FutureDate;
                                list[i].IsComplete = task.IsComplete;
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
                        if (AllTasks[k].Count == 0)
                            AllTasks.Remove(k);
                    }

                    // 3) If we didn't update in place, add to the target bucket (preserve existing relative ordering not possible here)
                    if (!updatedInPlace)
                    {
                        var copy = new TaskModel(task.TaskName, task.IsComplete, false, task.Description, new List<string>(task.People ?? new List<string>()), new List<string>(task.Meetings ?? new List<string>()), task.IsFuture, task.FutureDate, task.LinkPath, task.Id);
                        if (!AllTasks.ContainsKey(targetKey)) AllTasks[targetKey] = new List<TaskModel>();

                        // If there was an original stored entry we removed above, try to insert at its original index
                        var originalStored = storedEntries.FirstOrDefault();
                        if (originalStored != null && originalStored.Key != targetKey)
                        {
                            // If original bucket existed and we removed it, just append to target (cannot preserve cross-bucket position)
                            AllTasks[targetKey].Add(copy);
                        }
                        else if (originalStored == null)
                        {
                            // New task - append
                            AllTasks[targetKey].Add(copy);
                        }
                        else
                        {
                            // Shouldn't usually get here, but append as fallback
                            AllTasks[targetKey].Add(copy);
                        }

                        // If the task was visible in the current TaskList (we were viewing that date), remove it if it moved away
                        if (originKey == DateKey(_currentDate) && targetKey != originKey)
                        {
                            TaskList.Remove(task);
                            EnsureHasPlaceholder();
                        }
                    }
                    else
                    {
                        // If updated in place and we are viewing that date, update the UI model so it reflects any changes
                        if (_currentDate == DateTime.ParseExact(targetKey, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture))
                        {
                            // UI model 'task' already has the updated values from dialog above.
                            // Nothing more to do for ordering.
                        }
                        else
                        {
                            // If the task was updated in place in storage but we're not viewing that date,
                            // we don't need to modify TaskList here.
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

            IEnumerable<TaskModel> tasks = AllTasks.Values.SelectMany(list => list);

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

        private void ApplyPeopleFilter(string person)
        {
            TaskList.Clear();
            // Include date key so we can filter out old completed items when appropriate
            var entries = AllTasks.SelectMany(kv => kv.Value.Select(t => new { Task = t, DateKey = kv.Key }));
            var tasks = entries.Where(e => e.Task.People != null && e.Task.People.Count > 0).Select(e => new { e.Task, e.DateKey });
            if (!string.IsNullOrEmpty(person))
                tasks = tasks.Where(e => e.Task.People.Contains(person));

            bool showAll = PeopleShowAllToggle?.IsChecked == true;
            // Exclude tasks that are completed and dated yesterday or before unless 'show all' is set
            if (!showAll)
                tasks = tasks.Where(e => !IsOldCompleted(e.Task, e.DateKey));

            var unique = tasks.GroupBy(e => e.Task.Id).Select(g => g.First().Task).ToList();
            foreach (var t in unique)
            {
                TaskList.Add(new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id));
            }
        }

        private void ApplyMeetingsFilter(string meeting)
        {
            TaskList.Clear();
            var entries = AllTasks.SelectMany(kv => kv.Value.Select(t => new { Task = t, DateKey = kv.Key }));
            var tasks = entries.Where(e => e.Task.Meetings != null && e.Task.Meetings.Count > 0).Select(e => new { e.Task, e.DateKey });
            if (!string.IsNullOrEmpty(meeting))
                tasks = tasks.Where(e => e.Task.Meetings.Contains(meeting));

            bool showAll = MeetingsShowAllToggle?.IsChecked == true;
            // Exclude tasks that are completed and dated yesterday or before unless 'show all' is set
            if (!showAll)
                tasks = tasks.Where(e => !IsOldCompleted(e.Task, e.DateKey));

            var unique = tasks.GroupBy(e => e.Task.Id).Select(g => g.First().Task).ToList();
            foreach (var t in unique)
            {
                TaskList.Add(new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id));
            }
        }

        private void ApplyAllFilter(string term)
        {
            TaskList.Clear();

            // Build a sequence of tasks annotated with their source date key
            var entries = AllTasks.SelectMany(kv => kv.Value.Select(t => new { Task = t, DateKey = kv.Key })).ToList();

            if (!string.IsNullOrEmpty(term))
            {
                term = term.ToLowerInvariant();
                entries = entries.Where(e => (!string.IsNullOrEmpty(e.Task.TaskName) && e.Task.TaskName.ToLowerInvariant().Contains(term)) || (!string.IsNullOrEmpty(e.Task.Description) && e.Task.Description.ToLowerInvariant().Contains(term))).ToList();
            }

            // Deduplicate by Id, keeping the first occurrence
            var unique = entries.GroupBy(e => e.Task.Id).Select(g => g.First()).ToList();
            foreach (var e in unique)
            {
                var t = e.Task;
                var model = new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.LinkPath, t.Id);
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

                TaskList.Add(model);
            }
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
                            // open folder
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = path,
                                UseShellExecute = true,
                                Verb = "open"
                            });
                        }
                        else if (File.Exists(path))
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = path,
                                UseShellExecute = true,
                                Verb = "open"
                            });
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
                    var path = tm?.LinkPath;
                    if (string.IsNullOrWhiteSpace(path)) return;

                    try
                    {
                        // For folders and files, UseShellExecute=true will open with default handler
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = path,
                            UseShellExecute = true
                        });
                    }
                    catch
                    {
                        // If Process.Start fails for some reason, try URL handling
                        try
                        {
                            if (Uri.TryCreate(path, UriKind.Absolute, out var u))
                            {
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                {
                                    FileName = u.ToString(),
                                    UseShellExecute = true
                                });
                            }
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
                if (File.Exists(SaveFileName))
                {
                    File.Copy(SaveFileName, backupName, overwrite: true);
                }
                else
                {
                    // serialize current in-memory tasks to the backup file
                    var sanitized = new Dictionary<string, List<TaskModel>>();
                    foreach (var kv in AllTasks)
                         sanitized[kv.Key] = kv.Value.Where(t => t != null && !t.IsPlaceholder).ToList();
                    var json = JsonSerializer.Serialize(sanitized, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(backupName, json);
                }
                MessageBox.Show($"Backup saved: {backupName}", "Backup", MessageBoxButton.OK, MessageBoxImage.Information);
                UpdateLastBackupText(DateTime.Now);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Backup failed: {ex.Message}", "Backup", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Creates an automatic backup in the 'autobackups' folder and trims to the most recent AutoBackupKeep files.
        private void AutoBackupNow()
        {
            try
            {
                var folder = AutoBackupFolder;
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);

                var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                var backupName = $"{timestamp}_tasks_backup.json";
                var backupPath = System.IO.Path.Combine(folder, backupName);

                if (File.Exists(SaveFileName))
                {
                    File.Copy(SaveFileName, backupPath, overwrite: true);
                }
                else
                {
                    var sanitized = new Dictionary<string, List<TaskModel>>();
                    foreach (var kv in AllTasks)
                         sanitized[kv.Key] = kv.Value.Where(t => t != null && !t.IsPlaceholder).ToList();
                    var json = JsonSerializer.Serialize(sanitized, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(backupPath, json);
                }

                // Update UI with last backup time
                try { UpdateLastBackupText(DateTime.Now); } catch { }

                // Trim old backups leaving the most recent AutoBackupKeep files
                try
                {
                    var files = Directory.GetFiles(folder, "*_tasks_backup.json")
                        .OrderByDescending(f => System.IO.Path.GetFileName(f))
                        .ToList();
                    foreach (var f in files.Skip(AutoBackupKeep))
                    {
                        try { File.Delete(f); } catch { }
                    }
                }
                catch (Exception ex)
                {
                    // Log trimming errors but don't interrupt auto-backup
                    try { File.AppendAllText("tasks_debug_log.txt", $"[{DateTime.Now}] AutoBackup trim error: {ex.Message}\n"); } catch { }
                }
            }
            catch (Exception ex)
            {
                // Log errors silently for automatic backups
                try { File.AppendAllText("tasks_debug_log.txt", $"[{DateTime.Now}] AutoBackup error: {ex.Message}\n"); } catch { }
            }
        }

        // Find the latest backup timestamp from autobackups folder or root backups
        private DateTime? FindLatestBackupTimestamp()
        {
            try
            {
                var candidates = new List<DateTime>();

                // check autobackups
                if (Directory.Exists(AutoBackupFolder))
                {
                    foreach (var f in Directory.GetFiles(AutoBackupFolder, "*_tasks_backup.json"))
                    {
                        var name = System.IO.Path.GetFileNameWithoutExtension(f);
                        var parts = name.Split(new[] {"_tasks_backup"}, StringSplitOptions.None);
                        if (parts.Length > 0 && DateTime.TryParseExact(parts[0], "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var dt))
                            candidates.Add(dt);
                    }
                }

                // check root backups
                foreach (var f in Directory.GetFiles(".", "*_tasks_backup.json"))
                {
                    var name = System.IO.Path.GetFileNameWithoutExtension(f);
                    var parts = name.Split(new[] {"_tasks_backup"}, StringSplitOptions.None);
                    if (parts.Length > 0 && DateTime.TryParseExact(parts[0], "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var dt))
                        candidates.Add(dt);
                }

                if (candidates.Count == 0) return null;
                return candidates.Max();
            }
            catch { return null; }
        }

        private void UpdateLastBackupText(DateTime? dt)
        {
            try
            {
                if (LastBackupTextBlock == null) return;
                if (dt.HasValue)
                    LastBackupTextBlock.Text = $"Last backup: {dt.Value.ToString("yyyy-MM-dd HH:mm:ss")}";
                else
                    LastBackupTextBlock.Text = "Last backup: --";
            }
            catch { }
        }

        // Helper to determine whether a task is completed and belongs to yesterday or an earlier date
        private bool IsOldCompleted(TaskModel t, string dateKey)
        {
            if (t == null) return false;
            if (!t.IsComplete) return false;
            if (string.IsNullOrEmpty(dateKey)) return false;
            if (!DateTime.TryParseExact(dateKey, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var dt))
                return false;
            // Completed on or before yesterday should be hidden in People/Meetings modes
            return dt.Date <= DateTime.Today.AddDays(-1);
        }

        // Helper to populate People combobox according to toggle state
        private void PopulatePeopleComboBox()
        {
            try
            {
                bool showAll = PeopleShowAllToggle?.IsChecked == true;
                IEnumerable<string> people;
                // Build entries with source date so we can exclude old completed items when not showing all
                var entries = AllTasks.SelectMany(kv => kv.Value.Select(t => new { Task = t, DateKey = kv.Key }));
                if (showAll)
                {
                    // include all tasks (don't filter out old completed when 'show all' is on)
                    people = entries.SelectMany(e => e.Task.People ?? new List<string>()).Distinct();
                }
                else
                {
                    // only people who have active (incomplete) tasks (exclude old completed)
                    people = entries.Where(e => e.Task.People != null && e.Task.People.Count > 0 && !IsOldCompleted(e.Task, e.DateKey) && !e.Task.IsComplete)
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

        // Helper to populate Meetings combobox according to toggle state
        private void PopulateMeetingsComboBox()
        {
            try
            {
                bool showAll = MeetingsShowAllToggle?.IsChecked == true;
                IEnumerable<string> meetings;
                var entries = AllTasks.SelectMany(kv => kv.Value.Select(t => new { Task = t, DateKey = kv.Key }));
                if (showAll)
                {
                    // include all tasks (don't filter out old completed when 'show all' is on)
                    meetings = entries.SelectMany(e => e.Task.Meetings ?? new List<string>()).Distinct();
                }
                else
                {
                    // only meetings that have active (incomplete) tasks
                    meetings = entries.Where(e => e.Task.Meetings != null && e.Task.Meetings.Count > 0 && !IsOldCompleted(e.Task, e.DateKey) && !e.Task.IsComplete)
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

        // Toggle handlers for People
        private void PeopleShowAllToggle_Checked(object sender, RoutedEventArgs e)
        {
            PopulatePeopleComboBox();
            // Re-apply current selection (blank = all)
            if (PeopleFilterComboBox != null)
                ApplyPeopleFilter(PeopleFilterComboBox.SelectedItem as string);
        }

        private void PeopleShowAllToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            PopulatePeopleComboBox();
            if (PeopleFilterComboBox != null)
                ApplyPeopleFilter(PeopleFilterComboBox.SelectedItem as string);
        }

        // Toggle handlers for Meetings
        private void MeetingsShowAllToggle_Checked(object sender, RoutedEventArgs e)
        {
            PopulateMeetingsComboBox();
            if (MeetingsFilterComboBox != null)
                ApplyMeetingsFilter(MeetingsFilterComboBox.SelectedItem as string);
        }

        private void MeetingsShowAllToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            PopulateMeetingsComboBox();
            if (MeetingsFilterComboBox != null)
                ApplyMeetingsFilter(MeetingsFilterComboBox.SelectedItem as string);
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
    }
}

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
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

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

            Application.Current.Exit += Current_Exit;

            _isInitializing = false;
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
                                        var copy = new TaskModel(t.TaskName, false, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), true, item.FutureDate, Guid.NewGuid());
                                        if (!AllTasks.ContainsKey(newKey)) AllTasks[newKey] = new List<TaskModel>();
                                        AllTasks[newKey].Add(copy);
                                    }
                                    else // Copy to today (default)
                                    {
                                        var todayKey = DateKey(today);
                                        var copy = new TaskModel(t.TaskName, false, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), false, null, Guid.NewGuid());
                                        if (!AllTasks.ContainsKey(todayKey)) AllTasks[todayKey] = new List<TaskModel>();
                                        AllTasks[todayKey].Add(copy);
                                    }
                                }

                                SaveTasks();

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
            SaveCurrentDateTasks();
            SaveTasks();
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
                SaveCurrentDateTasks(); // Ensure current day's tasks are saved before serializing
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
            // SaveCurrentDateTasks(); // Removed to prevent wiping out tasks on startup
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
                    var model = new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.Id);
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

        // In TaskTextBox_PreviewTextInput, prevent typing for non-today
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

                // update current date storage
                var key = DateKey(_currentDate);
                AllTasks[key] = TaskList.Where(t => !t.IsPlaceholder)
                    .Select(t => new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate))
                    .ToList();
                SaveTasks(); // Save immediately after any change for today
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
            // Save after key up for today
            SaveCurrentDateTasks();
            SaveTasks();
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
            // Save after lost focus for today
            SaveCurrentDateTasks();
            SaveTasks();
        }

        // Prevent checking/unchecking for non-today
        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (_currentDate != DateTime.Today)
            {
                if (sender is CheckBox cb)
                {
                    cb.IsChecked = ((TaskModel)cb.DataContext).IsComplete;
                }
                e.Handled = true;
            }
            else
            {
                // Save after completion state changes for today
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

        private void TodayButton_Click(object sender, RoutedEventArgs e)
        {
            // Exit any special filter mode and return to today's view
            // restore date controls
            DateControlsGrid.Visibility = Visibility.Visible;
            // hide people/meetings filter comboboxes
            if (PeopleFilterComboBox != null) PeopleFilterComboBox.Visibility = Visibility.Collapsed;
            if (MeetingsFilterComboBox != null) MeetingsFilterComboBox.Visibility = Visibility.Collapsed;
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
            // hide date controls and show people combobox
            DateControlsGrid.Visibility = Visibility.Collapsed;
            PeopleFilterComboBox.Visibility = Visibility.Visible;
            // hide search box if visible
            if (SearchTextBox != null) SearchTextBox.Visibility = Visibility.Collapsed;
            if (SearchBoxBorder != null) SearchBoxBorder.Visibility = Visibility.Collapsed;
            MeetingsFilterComboBox.Visibility = Visibility.Collapsed;
            // populate people combobox
            var people = AllTasks.Values.SelectMany(l => l.SelectMany(t => t.People)).Distinct().OrderBy(s => s).ToList();
            people.Insert(0, ""); // blank = any
            PeopleFilterComboBox.ItemsSource = people;
            PeopleFilterComboBox.SelectedIndex = 0;
            // show all tasks that have any people
            // ApplyPeopleFilter may be implemented elsewhere; if not, fallback to ApplyFilter
            try { ApplyPeopleFilter(null); } catch { ApplyFilter(); }
            UpdateTitle();
        }

        private void MeetingButton_Click(object sender, RoutedEventArgs e)
        {
            _mode = ViewMode.Meetings;
            // hide date controls and show meetings combobox
            DateControlsGrid.Visibility = Visibility.Collapsed;
            MeetingsFilterComboBox.Visibility = Visibility.Visible;
            // hide search box if visible
            if (SearchTextBox != null) SearchTextBox.Visibility = Visibility.Collapsed;
            if (SearchBoxBorder != null) SearchBoxBorder.Visibility = Visibility.Collapsed;
            PeopleFilterComboBox.Visibility = Visibility.Collapsed;
            // populate meetings combobox
            var meetings = AllTasks.Values.SelectMany(l => l.SelectMany(t => t.Meetings)).Distinct().OrderBy(s => s).ToList();
            meetings.Insert(0, ""); // blank = any
            MeetingsFilterComboBox.ItemsSource = meetings;
            MeetingsFilterComboBox.SelectedIndex = 0;
            // show all tasks that have any meetings
            try { ApplyMeetingsFilter(null); } catch { ApplyFilter(); }
            UpdateTitle();
        }

        private void AllButton_Click(object sender, RoutedEventArgs e)
        {
            _mode = ViewMode.All;
            FilterPopup.IsOpen = false;
            CalendarPopup.IsOpen = false;

            // Show search box and hide date controls
            DateControlsGrid.Visibility = Visibility.Collapsed;
            PeopleFilterComboBox.Visibility = Visibility.Collapsed;
            MeetingsFilterComboBox.Visibility = Visibility.Collapsed;
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

        private void ExitAllMode()
        {
            DateControlsGrid.Visibility = Visibility.Visible;
            if (SearchTextBox != null) SearchTextBox.Visibility = Visibility.Collapsed;
            if (SearchBoxBorder != null) SearchBoxBorder.Visibility = Visibility.Collapsed;
            _mode = ViewMode.Today;
            SetCurrentDate(_currentDate);
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
                .Select(t => new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.Id))
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
                    // update task properties from dialog
                    task.TaskName = dialog.TaskTitle;
                    task.Description = dialog.TaskDescription;
                    task.People = new List<string>(dialog.TaskPeople);
                    task.Meetings = new List<string>(dialog.TaskMeetings);

                    // handle future scheduling
                    task.IsFuture = dialog.IsFuture;
                    task.FutureDate = dialog.FutureDate;

                    var todayKey = DateKey(DateTime.Today);

                    if (task.IsFuture && task.FutureDate.HasValue)
                    {
                        var newKey = DateKey(task.FutureDate.Value.Date);

                        // remove this task from any date it currently exists under (by id)
                        foreach (var k in AllTasks.Keys.ToList())
                        {
                            AllTasks[k].RemoveAll(t => t.Id == originalId);
                            if (AllTasks[k].Count == 0) AllTasks.Remove(k);
                        }

                        // add to AllTasks under future date
                        var copy = new TaskModel(task.TaskName, task.IsComplete, false, task.Description, new List<string>(task.People), new List<string>(task.Meetings), true, task.FutureDate, task.Id);
                        if (!AllTasks.ContainsKey(newKey)) AllTasks[newKey] = new List<TaskModel>();
                        AllTasks[newKey].Add(copy);

                        // remove from current TaskList if viewing that date
                        if (originKey == DateKey(_currentDate))
                        {
                            TaskList.Remove(task);
                            EnsureHasPlaceholder();
                        }
                    }
                    else
                    {
                        // remove this task from any date it currentlyexists under
                        foreach (var k in AllTasks.Keys.ToList())
                        {
                            AllTasks[k].RemoveAll(t => t.Id == originalId);
                            if (AllTasks[k].Count == 0) AllTasks.Remove(k);
                        }

                        // move to today's storage
                        var copy = new TaskModel(task.TaskName, task.IsComplete, false, task.Description, new List<string>(task.People), new List<string>(task.Meetings), false, null, task.Id);
                        if (!AllTasks.ContainsKey(todayKey)) AllTasks[todayKey] = new List<TaskModel>();
                        AllTasks[todayKey].Add(copy);

                        // ensure TaskList contains the task instance (moved back to today)
                        if (_currentDate == DateTime.Today)
                        {
                            if (!TaskList.Contains(task))
                            {
                                task.IsReadOnly = false;
                                TaskList.Insert(TaskList.Count - 1, task);
                            }
                        }
                    }

                    SaveTasks(); // Save after editing a task
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
                TaskList.Add(new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.Id));
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
            var tasks = AllTasks.Values.SelectMany(list => list).Where(t => t.People != null && t.People.Count > 0);
            if (!string.IsNullOrEmpty(person))
                tasks = tasks.Where(t => t.People.Contains(person));

            var unique = tasks.GroupBy(t => t.Id).Select(g => g.First()).ToList();
            foreach (var t in unique)
            {
                TaskList.Add(new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.Id));
            }
        }

        private void ApplyMeetingsFilter(string meeting)
        {
            TaskList.Clear();
            var tasks = AllTasks.Values.SelectMany(list => list).Where(t => t.Meetings != null && t.Meetings.Count > 0);
            if (!string.IsNullOrEmpty(meeting))
                tasks = tasks.Where(t => t.Meetings.Contains(meeting));

            var unique = tasks.GroupBy(t => t.Id).Select(g => g.First()).ToList();
            foreach (var t in unique)
            {
                TaskList.Add(new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.Id));
            }
        }

        private void ApplyAllFilter(string term)
        {
            TaskList.Clear();
            var tasks = AllTasks.Values.SelectMany(list => list).ToList();
            if (!string.IsNullOrEmpty(term))
            {
                term = term.ToLowerInvariant();
                tasks = tasks.Where(t => (!string.IsNullOrEmpty(t.TaskName) && t.TaskName.ToLowerInvariant().Contains(term)) || (!string.IsNullOrEmpty(t.Description) && t.Description.ToLowerInvariant().Contains(term))).ToList();
            }

            var unique = tasks.GroupBy(t => t.Id).Select(g => g.First()).ToList();
            foreach (var t in unique)
            {
                TaskList.Add(new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings), t.IsFuture, t.FutureDate, t.Id));
            }
        }
    }
}

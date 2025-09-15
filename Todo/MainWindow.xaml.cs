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

        // All tasks keyed by date string (yyyy-MM-dd)
        private Dictionary<string, List<TaskModel>> AllTasks { get; set; } = new Dictionary<string, List<TaskModel>>();

        public ObservableCollection<TaskModel> TaskList { get; set; } = new ObservableCollection<TaskModel>();

        private DateTime _currentDate = DateTime.Today;

        public MainWindow()
        {
            InitializeComponent();

            DataContext = this;

            LoadTasks();

            SetCurrentDate(_currentDate);

            lbTasksList.ItemsSource = TaskList;

            Application.Current.Exit += Current_Exit;
        }

        private void Current_Exit(object sender, ExitEventArgs e)
        {
            SaveCurrentDateTasks();
            SaveTasks();
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
                        AllTasks = items;
                    }
                }
            }
            catch
            {
                // ignore load errors
            }
        }

        private void SaveTasks()
        {
            try
            {
                SaveCurrentDateTasks(); // Ensure current day's tasks are saved before serializing
                // ensure we don't save placeholder items
                var sanitized = new Dictionary<string, List<TaskModel>>();
                foreach (var kv in AllTasks)
                {
                    sanitized[kv.Key] = kv.Value.Where(t => !t.IsPlaceholder).ToList();
                }

                var json = JsonSerializer.Serialize(sanitized, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SaveFileName, json);
            }
            catch
            {
                // ignore save errors
            }
        }

        private string DateKey(DateTime dt) => dt.ToString("yyyy-MM-dd");

        private void SetCurrentDate(DateTime dt)
        {
            SaveCurrentDateTasks();
            _currentDate = dt.Date;
            CurrentDayButton.Content = FormatDate(_currentDate);
            LoadTasksForDate(_currentDate);
            bool isToday = _currentDate == DateTime.Today;
            foreach (var task in TaskList)
            {
                task.IsReadOnly = !isToday;
            }
        }

        private void LoadTasksForDate(DateTime dt)
        {
            TaskList.Clear();
            var key = DateKey(dt);
            if (AllTasks.ContainsKey(key))
            {
                foreach (var t in AllTasks[key])
                {
                    // Deep copy all properties
                    var model = new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings));
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
                AllTasks[key] = TaskList.Where(t => !t.IsPlaceholder).Select(t => new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings))).ToList();
                SaveTasks(); // Save immediately after any change for today
            }
            _placeholderJustFocused = false;
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

        // In TaskTextBox_KeyUp, prevent key for non-today
        private void TaskTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (_currentDate != DateTime.Today)
            {
                e.Handled = true;
                return;
            }
        }

        // In TaskTextBox_LostFocus, do nothing for non-today
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

        private void CalendarControl_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CalendarControl.SelectedDate.HasValue)
            {
                SetCurrentDate(CalendarControl.SelectedDate.Value);
                CalendarPopup.IsOpen = false;
            }
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
            AllTasks[key] = TaskList.Where(t => !t.IsPlaceholder)
                .Select(t => new TaskModel(t.TaskName, t.IsComplete, false, t.Description, new List<string>(t.People), new List<string>(t.Meetings)))
                .ToList();
        }

        private void TaskOptionsButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is TaskModel task && !task.IsPlaceholder)
            {
                // Gather suggestions from all tasks
                var allPeople = AllTasks.Values.SelectMany(list => list.SelectMany(t => t.People)).Distinct().ToList();
                var allMeetings = AllTasks.Values.SelectMany(list => list.SelectMany(t => t.Meetings)).Distinct().ToList();
                var dialog = new TaskDialog(task, allPeople, allMeetings) { Owner = this };
                if (dialog.ShowDialog() == true)
                {
                    task.TaskName = dialog.TaskTitle;
                    task.Description = dialog.TaskDescription;
                    task.People = new List<string>(dialog.TaskPeople);
                    task.Meetings = new List<string>(dialog.TaskMeetings);
                    SaveTasks(); // Save after editing a task
                }
            }
        }
    }
}

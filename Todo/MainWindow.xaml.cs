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
        private const string SaveFileName = "tasks.json";

        public ObservableCollection<TaskModel> TaskList { get; set; } = new ObservableCollection<TaskModel>();

        public MainWindow()
        {
            InitializeComponent();

            DataContext = this;

            LoadTasks();

            if (TaskList.Count == 0)
            {
                SetupSampleTasks();
            }

            EnsureHasPlaceholder();

            lbTasksList.ItemsSource = TaskList;

            Application.Current.Exit += Current_Exit;
        }

        private void Current_Exit(object sender, ExitEventArgs e)
        {
            SaveTasks();
        }

        private void LoadTasks()
        {
            try
            {
                if (File.Exists(SaveFileName))
                {
                    var json = File.ReadAllText(SaveFileName);
                    var items = JsonSerializer.Deserialize<List<TaskModel>>(json);
                    if (items != null)
                    {
                        TaskList.Clear();
                        foreach (var it in items)
                        {
                            // keep IsPlaceholder false on load; we'll re-add placeholder below
                            it.IsPlaceholder = false;
                            TaskList.Add(it);
                        }
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
                // do not persist placeholder item
                var toSave = TaskList.Where(t => !t.IsPlaceholder).ToList();
                var json = JsonSerializer.Serialize(toSave, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SaveFileName, json);
            }
            catch
            {
                // ignore save errors
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

        private void LoadSampleTasksToScreen()
        {
            lbTasksList.ItemsSource = null;
            lbTasksList.ItemsSource = TaskList;
        }

        private void SetupSampleTasks()
        {
            TaskList.Add(new TaskModel("Task 1"));
            TaskList.Add(new TaskModel("Task 2"));
            TaskList.Add(new TaskModel("Task 3"));
            TaskList.Add(new TaskModel("Task 4"));
            TaskList.Add(new TaskModel("Task 5"));
            // placeholder
            TaskList.Add(new TaskModel(string.Empty, false, true));
        }

        private void ListViewItem_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

        }

        private bool _suppressTextChanged = false;
        private bool _placeholderJustFocused = false;

        private void TaskTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                if (tm.IsPlaceholder)
                {
                    // Don't select text, just set a flag
                    _placeholderJustFocused = true;
                }
            }
        }

        private void TaskTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                if (tm.IsPlaceholder)
                {
                    // As soon as user types, clear the placeholder and set to normal
                    _suppressTextChanged = true;
                    tm.IsPlaceholder = false;
                    tm.TaskName = "";
                    tb.Text = "";
                    _suppressTextChanged = false;
                    // Let the input go through
                }
            }
        }

        private void TaskTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressTextChanged) return;
            if (sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                // Always add a new placeholder as soon as the last line's text changes from empty to non-empty
                if (!tm.IsPlaceholder && TaskList.Last() == tm && !string.IsNullOrWhiteSpace(tb.Text) && TaskList.Count(x => x.IsPlaceholder) == 0)
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
            }
            _placeholderJustFocused = false;
        }

        private void TaskTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            // No longer used for placeholder logic
        }

        private void TaskTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            // if placeholder and empty, restore placeholder text
            if (sender is TextBox tb && tb.DataContext is TaskModel tm)
            {
                if (string.IsNullOrWhiteSpace(tm.TaskName))
                {
                    if (!tm.IsPlaceholder)
                    {
                        // convert to placeholder
                        tm.TaskName = "Type new task here...";
                        tm.IsPlaceholder = true;
                    }
                }
            }
        }
    }
}

using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Collections.ObjectModel;
using System.Windows.Input;
using System;
using System.ComponentModel;

namespace Todo
{
    public partial class TaskDialog : Window, INotifyPropertyChanged
    {
        public string TaskTitle { get; set; }
        public string TaskDescription { get; set; }
        public ObservableCollection<string> TaskPeople { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> TaskMeetings { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> PeopleSuggestions { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> MeetingSuggestions { get; set; } = new ObservableCollection<string>();
        private string _linkPath = string.Empty;
        public string LinkPath
        {
            get => _linkPath;
            set
            {
                if (_linkPath == value) return;
                _linkPath = value ?? string.Empty;
                // Keep the displayed textbox in sync if available
                try { if (LinkBox != null) LinkBox.Text = _linkPath; } catch { }
                OnPropertyChanged(nameof(LinkPath));
                UpdateLinkFlags();
            }
        }

        private bool _isHyperlink;
        public bool IsHyperlink
        {
            get => _isHyperlink;
            private set
            {
                if (_isHyperlink == value) return;
                _isHyperlink = value;
                OnPropertyChanged(nameof(IsHyperlink));
            }
        }

        private bool _hasLink;
        public bool HasLink
        {
            get => _hasLink;
            private set
            {
                if (_hasLink == value) return;
                _hasLink = value;
                OnPropertyChanged(nameof(HasLink));
            }
        }

        private bool _isFuture;
        public bool IsFuture
        {
            get => _isFuture;
            set
            {
                if (_isFuture == value) return;
                _isFuture = value;
                OnPropertyChanged(nameof(IsFuture));
            }
        }

        private DateTime? _futureDate;
        public DateTime? FutureDate
        {
            get => _futureDate;
            set
            {
                if (_futureDate == value) return;
                _futureDate = value;
                OnPropertyChanged(nameof(FutureDate));
            }
        }

        public TaskDialog(TaskModel model, IEnumerable<string> peopleSuggestions, IEnumerable<string> meetingSuggestions)
        {
            InitializeComponent();
            TaskTitle = model.TaskName;
            TaskDescription = model.Description;
            TaskPeople = new ObservableCollection<string>(model.People ?? new List<string>());
            TaskMeetings = new ObservableCollection<string>(model.Meetings ?? new List<string>());
            PeopleSuggestions = new ObservableCollection<string>(peopleSuggestions ?? Enumerable.Empty<string>());
            MeetingSuggestions = new ObservableCollection<string>(meetingSuggestions ?? Enumerable.Empty<string>());
            IsFuture = model.IsFuture;
            FutureDate = model.FutureDate;

            TitleBox.Text = TaskTitle;
            DescriptionBox.Text = TaskDescription;
            PeopleList.ItemsSource = TaskPeople;
            MeetingsList.ItemsSource = TaskMeetings;
            PeopleBox.ItemsSource = PeopleSuggestions;
            MeetingsBox.ItemsSource = MeetingSuggestions;
            DataContext = this;

            // wire up handlers
            PeopleBox.KeyDown += PeopleBox_KeyDown;
            MeetingsBox.KeyDown += MeetingsBox_KeyDown;
            PeopleList.MouseDoubleClick += PeopleList_MouseDoubleClick;
            MeetingsList.MouseDoubleClick += MeetingsList_MouseDoubleClick;

            // initialize link
            LinkBox.Text = model.LinkPath ?? string.Empty;
            LinkPath = model.LinkPath ?? string.Empty;
        }

        private void PeopleBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddPersonFromInput();
                e.Handled = true;
            }
        }

        private void MeetingsBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddMeetingFromInput();
                e.Handled = true;
            }
        }

        private void PeopleList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (PeopleList.SelectedItem is string s)
            {
                TaskPeople.Remove(s);
            }
        }

        private void MeetingsList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (MeetingsList.SelectedItem is string s)
            {
                TaskMeetings.Remove(s);
            }
        }

        private void AddPersonFromInput()
        {
            var text = (PeopleBox.Text ?? string.Empty).Trim();
            if (!string.IsNullOrEmpty(text) && !TaskPeople.Contains(text))
            {
                TaskPeople.Add(text);
                if (!PeopleSuggestions.Contains(text)) PeopleSuggestions.Add(text);
            }
            PeopleBox.Text = string.Empty;
        }

        private void AddMeetingFromInput()
        {
            var text = (MeetingsBox.Text ?? string.Empty).Trim();
            if (!string.IsNullOrEmpty(text) && !TaskMeetings.Contains(text))
            {
                TaskMeetings.Add(text);
                if (!MeetingSuggestions.Contains(text)) MeetingSuggestions.Add(text);
            }
            MeetingsBox.Text = string.Empty;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            TaskTitle = TitleBox.Text;
            TaskDescription = DescriptionBox.Text;
            // capture link
            LinkPath = LinkBox.Text;
            // ensure Future properties are updated from UI bindings
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void AddPersonButton_Click(object sender, RoutedEventArgs e)
        {
            AddPersonFromInput();
        }

        private void AddMeetingButton_Click(object sender, RoutedEventArgs e)
        {
            AddMeetingFromInput();
        }

        private void BrowseLinkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new Microsoft.Win32.OpenFileDialog();
                if (dlg.ShowDialog(this) == true)
                {
                    LinkPath = dlg.FileName;
                }
            }
            catch { }
        }

        private void BrowseFolderButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var fbd = new System.Windows.Forms.FolderBrowserDialog())
                {
                    var res = fbd.ShowDialog();
                    if (res == System.Windows.Forms.DialogResult.OK)
                    {
                        LinkPath = fbd.SelectedPath;
                    }
                }
            }
            catch { }
        }

        private void PasteLinkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Clipboard.ContainsText())
                {
                    var text = (Clipboard.GetText() ?? string.Empty).Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        LinkPath = text;
                    }
                    else
                    {
                        MessageBox.Show("Clipboard does not contain a valid link.", "Paste Link", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Clipboard does not contain text to paste.", "Paste Link", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch { }
        }

        private void UpdateLinkFlags()
        {
            HasLink = !string.IsNullOrWhiteSpace(_linkPath);
            bool isHttp = false;
            if (HasLink)
            {
                try
                {
                    if (Uri.TryCreate(_linkPath, UriKind.Absolute, out var u))
                    {
                        var s = u.Scheme?.ToLowerInvariant();
                        if (s == "http" || s == "https") isHttp = true;
                      }
                }
                catch { }
            }
            IsHyperlink = isHttp;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

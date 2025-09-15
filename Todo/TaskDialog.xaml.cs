using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Todo
{
    public partial class TaskDialog : Window
    {
        public string TaskTitle { get; set; }
        public string TaskDescription { get; set; }
        public List<string> TaskPeople { get; set; } = new List<string>();
        public List<string> TaskMeetings { get; set; } = new List<string>();
        public List<string> PeopleSuggestions { get; set; } = new List<string>();
        public List<string> MeetingSuggestions { get; set; } = new List<string>();

        public TaskDialog(TaskModel model, IEnumerable<string> peopleSuggestions, IEnumerable<string> meetingSuggestions)
        {
            InitializeComponent();
            TaskTitle = model.TaskName;
            TaskDescription = model.Description;
            TaskPeople = new List<string>(model.People);
            TaskMeetings = new List<string>(model.Meetings);
            PeopleSuggestions = peopleSuggestions.ToList();
            MeetingSuggestions = meetingSuggestions.ToList();
            TitleBox.Text = TaskTitle;
            DescriptionBox.Text = TaskDescription;
            PeopleList.ItemsSource = TaskPeople;
            MeetingsList.ItemsSource = TaskMeetings;
            DataContext = this;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            TaskTitle = TitleBox.Text;
            TaskDescription = DescriptionBox.Text;
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

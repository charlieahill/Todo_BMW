using System;
using System.Linq;
using System.Windows;

namespace Todo
{
    public partial class TimeTemplateDialog : Window
    {
        public TimeTemplate Template { get; private set; }

        public TimeTemplateDialog()
        {
            InitializeComponent();
            Template = new TimeTemplate();
            StartDatePicker.SelectedDate = Template.StartDate;
            EndDatePicker.SelectedDate = Template.EndDate;
            JobText.Text = Template.JobDescription;
            Mon.Text = Template.HoursPerWeekday[0].ToString();
            Tue.Text = Template.HoursPerWeekday[1].ToString();
            Wed.Text = Template.HoursPerWeekday[2].ToString();
            Thu.Text = Template.HoursPerWeekday[3].ToString();
            Fri.Text = Template.HoursPerWeekday[4].ToString();
            Sat.Text = Template.HoursPerWeekday[5].ToString();
            Sun.Text = Template.HoursPerWeekday[6].ToString();
        }

        public TimeTemplateDialog(TimeTemplate t) : this()
        {
            Template = t;
            StartDatePicker.SelectedDate = t.StartDate;
            EndDatePicker.SelectedDate = t.EndDate;
            JobText.Text = t.JobDescription;
            Mon.Text = t.HoursPerWeekday[0].ToString();
            Tue.Text = t.HoursPerWeekday[1].ToString();
            Wed.Text = t.HoursPerWeekday[2].ToString();
            Thu.Text = t.HoursPerWeekday[3].ToString();
            Fri.Text = t.HoursPerWeekday[4].ToString();
            Sat.Text = t.HoursPerWeekday[5].ToString();
            Sun.Text = t.HoursPerWeekday[6].ToString();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            Template.StartDate = StartDatePicker.SelectedDate ?? DateTime.Today;
            Template.EndDate = EndDatePicker.SelectedDate;
            Template.JobDescription = JobText.Text ?? "";
            double[] arr = new double[7];
            double.TryParse(Mon.Text, out arr[0]);
            double.TryParse(Tue.Text, out arr[1]);
            double.TryParse(Wed.Text, out arr[2]);
            double.TryParse(Thu.Text, out arr[3]);
            double.TryParse(Fri.Text, out arr[4]);
            double.TryParse(Sat.Text, out arr[5]);
            double.TryParse(Sun.Text, out arr[6]);
            Template.HoursPerWeekday = arr;
            this.DialogResult = true;
            this.Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}

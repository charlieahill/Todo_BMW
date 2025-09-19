using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace Todo
{
    public partial class TimeTrackingDialog : Window
    {
        public TimeTrackingDialog()
        {
            InitializeComponent();

            var svc = TimeTrackingService.Instance;
            ToDatePicker.SelectedDate = DateTime.Today;
            FromDatePicker.SelectedDate = DateTime.Today.AddDays(-13);

            LoadTemplates();
            RefreshSummaries();
        }

        private void LoadTemplates()
        {
            TemplatesList.ItemsSource = TimeTrackingService.Instance.GetTemplates();
        }

        private void RefreshSummaries()
        {
            var from = FromDatePicker.SelectedDate ?? DateTime.Today.AddDays(-13);
            var to = ToDatePicker.SelectedDate ?? DateTime.Today;
            var days = TimeTrackingService.Instance.GetDaySummaries(from, to);
            SummariesGrid.ItemsSource = days;

            TotalWorkedText.Text = days.Where(d => d.WorkedHours.HasValue).Sum(d => d.WorkedHours.Value).ToString("0.00");
            TotalStandardText.Text = days.Sum(d => d.StandardHours).ToString("0.00");
        }

        private void DateRange_Changed(object sender, SelectionChangedEventArgs e)
        {
            RefreshSummaries();
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshSummaries();
        }

        private void AddTemplate_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new TimeTemplateDialog() { Owner = this };
            if (dlg.ShowDialog() == true)
            {
                TimeTrackingService.Instance.UpsertTemplate(dlg.Template);
                LoadTemplates();
            }
        }

        private void EditTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (TemplatesList.SelectedItem is TimeTemplate t)
            {
                var copy = new TimeTemplate
                {
                    Id = t.Id,
                    StartDate = t.StartDate,
                    EndDate = t.EndDate,
                    JobDescription = t.JobDescription,
                    Location = t.Location,
                    HoursPerWeekday = (double[])t.HoursPerWeekday.Clone()
                };
                var dlg = new TimeTemplateDialog(copy) { Owner = this };
                if (dlg.ShowDialog() == true)
                {
                    TimeTrackingService.Instance.UpsertTemplate(dlg.Template);
                    LoadTemplates();
                }
            }
        }

        private void TemplatesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            EditTemplateButton.IsEnabled = TemplatesList.SelectedItem != null;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}

using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Text;

namespace Todo
{
    public partial class TimeTemplateDialog : Window
    {
        public TimeTemplate Template { get; private set; }

        public TimeTemplateDialog()
        {
            InitializeComponent();
            Template = new TimeTemplate();

            // Default template name as requested
            Template.TemplateName = "Employment Position (Location)";

            StartDatePicker.SelectedDate = Template.StartDate;
            EndDatePicker.SelectedDate = Template.EndDate;
            TemplateNameText.Text = Template.TemplateName;
            EmploymentPositionText.Text = Template.EmploymentPosition;
            EmploymentLocationText.Text = Template.EmploymentLocation;

            Mon.Text = Template.HoursPerWeekday[0].ToString();
            Tue.Text = Template.HoursPerWeekday[1].ToString();
            Wed.Text = Template.HoursPerWeekday[2].ToString();
            Thu.Text = Template.HoursPerWeekday[3].ToString();
            Fri.Text = Template.HoursPerWeekday[4].ToString();
            Sat.Text = Template.HoursPerWeekday[5].ToString();
            Sun.Text = Template.HoursPerWeekday[6].ToString();

            OngoingCheck.Checked += (s, e) => EndDatePicker.IsEnabled = false;
            OngoingCheck.Unchecked += (s, e) => EndDatePicker.IsEnabled = true;
            OngoingCheck.IsChecked = Template.EndDate == null;
            EndDatePicker.IsEnabled = !(OngoingCheck.IsChecked == true);

            // New fields: standard times and lunch (lunch now in HH:MM)
            StandardStartBox.Text = Template.StandardStart.ToString(@"hh\:mm");
            LunchBreakBox.Text = Template.LunchBreak.ToString(@"hh\:mm");
            StandardEndBox.Text = Template.StandardEnd.ToString(@"hh\:mm");
        }

        public TimeTemplateDialog(TimeTemplate t) : this()
        {
            if (t == null) throw new ArgumentNullException(nameof(t));
            Template = t;
            StartDatePicker.SelectedDate = t.StartDate;
            EndDatePicker.SelectedDate = t.EndDate;
            TemplateNameText.Text = t.TemplateName ?? t.JobDescription ?? string.Empty;
            EmploymentPositionText.Text = t.EmploymentPosition ?? string.Empty;
            EmploymentLocationText.Text = t.EmploymentLocation ?? string.Empty;
            if (t.HoursPerWeekday != null && t.HoursPerWeekday.Length == 7)
            {
                Mon.Text = t.HoursPerWeekday[0].ToString();
                Tue.Text = t.HoursPerWeekday[1].ToString();
                Wed.Text = t.HoursPerWeekday[2].ToString();
                Thu.Text = t.HoursPerWeekday[3].ToString();
                Fri.Text = t.HoursPerWeekday[4].ToString();
                Sat.Text = t.HoursPerWeekday[5].ToString();
                Sun.Text = t.HoursPerWeekday[6].ToString();
            }
            OngoingCheck.IsChecked = t.EndDate == null;
            EndDatePicker.IsEnabled = !(OngoingCheck.IsChecked == true);

            // New fields
            StandardStartBox.Text = t.StandardStart.ToString(@"hh\:mm");
            LunchBreakBox.Text = t.LunchBreak.ToString(@"hh\:mm");
            StandardEndBox.Text = t.StandardEnd.ToString(@"hh\:mm");
        }

        private bool ValidateInputs(out string msg)
        {
            msg = null;
            if (string.IsNullOrWhiteSpace(TemplateNameText.Text)) { msg = "Please enter a template name."; return false; }
            if (!(OngoingCheck.IsChecked == true))
            {
                if (EndDatePicker.SelectedDate.HasValue && StartDatePicker.SelectedDate.HasValue && EndDatePicker.SelectedDate.Value.Date < StartDatePicker.SelectedDate.Value.Date)
                {
                    msg = "End date must be the same or after start date."; return false;
                }
            }
            double v;
            TextBox[] boxes = new TextBox[] { Mon, Tue, Wed, Thu, Fri, Sat, Sun };
            for (int i = 0; i < boxes.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(boxes[i].Text)) boxes[i].Text = "0";
                if (!double.TryParse(boxes[i].Text, out v) || v < 0 || v > 24) { msg = "Please enter numeric hours between 0 and 24 for each weekday."; return false; }
            }

            // validate new fields (times in HH:MM)
            if (string.IsNullOrWhiteSpace(StandardStartBox.Text) || !TimeSpan.TryParse(StandardStartBox.Text, out var _)) { msg = "Please enter a valid standard start time (HH:MM)."; return false; }
            if (string.IsNullOrWhiteSpace(StandardEndBox.Text) || !TimeSpan.TryParse(StandardEndBox.Text, out var _)) { msg = "Please enter a valid standard end time (HH:MM)."; return false; }
            if (string.IsNullOrWhiteSpace(LunchBreakBox.Text) || !TimeSpan.TryParse(LunchBreakBox.Text, out var lunch) || lunch < TimeSpan.Zero || lunch.TotalHours > 24) { msg = "Please enter a valid lunch break (HH:MM)."; return false; }

            return true;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInputs(out var err)) { MessageBox.Show(err, "Validation", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            Template.StartDate = StartDatePicker.SelectedDate ?? DateTime.Today;
            Template.EndDate = (OngoingCheck.IsChecked == true) ? (DateTime?)null : EndDatePicker.SelectedDate;
            Template.TemplateName = TemplateNameText.Text ?? string.Empty;
            Template.EmploymentPosition = EmploymentPositionText.Text ?? string.Empty;
            Template.EmploymentLocation = EmploymentLocationText.Text ?? string.Empty;
            double[] arr = new double[7];
            double.TryParse(Mon.Text, out arr[0]);
            double.TryParse(Tue.Text, out arr[1]);
            double.TryParse(Wed.Text, out arr[2]);
            double.TryParse(Thu.Text, out arr[3]);
            double.TryParse(Fri.Text, out arr[4]);
            double.TryParse(Sat.Text, out arr[5]);
            double.TryParse(Sun.Text, out arr[6]);
            Template.HoursPerWeekday = arr;

            // new fields
            Template.StandardStart = TimeSpan.Parse(StandardStartBox.Text);
            Template.LunchBreak = TimeSpan.Parse(LunchBreakBox.Text);
            Template.StandardEnd = TimeSpan.Parse(StandardEndBox.Text);

            this.DialogResult = true;
            this.Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        public static TimeTemplate ShowDialogForTemplate(Window owner)
        {
            var dlg = new TimeTemplateDialog() { Owner = owner };
            if (dlg.ShowDialog() == true)
                return dlg.Template;
            return null;
        }

        // When standard times change, update default weekday hours (finish - start - lunch)
        private void StandardTimes_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!TimeSpan.TryParse(StandardStartBox.Text, out var start)) return;
            if (!TimeSpan.TryParse(StandardEndBox.Text, out var end)) return;
            if (!TimeSpan.TryParse(LunchBreakBox.Text, out var lunch)) return;

            var worked = (end - start - lunch).TotalHours;
            if (worked < 0) worked = 0;
            var s = worked.ToString("0.##");

            // Update Mon..Fri defaults; leave Sat/Sun as-is
            Mon.Text = s;
            Tue.Text = s;
            Wed.Text = s;
            Thu.Text = s;
            Fri.Text = s;
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // ensure template object is current
                if (!ValidateInputs(out var err)) { MessageBox.Show(err, "Validation", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
                Template.StartDate = StartDatePicker.SelectedDate ?? DateTime.Today;
                Template.EndDate = (OngoingCheck.IsChecked == true) ? (DateTime?)null : EndDatePicker.SelectedDate;
                Template.TemplateName = TemplateNameText.Text ?? string.Empty;
                Template.EmploymentPosition = EmploymentPositionText.Text ?? string.Empty;
                Template.EmploymentLocation = EmploymentLocationText.Text ?? string.Empty;
                double[] arr = new double[7];
                double.TryParse(Mon.Text, out arr[0]);
                double.TryParse(Tue.Text, out arr[1]);
                double.TryParse(Wed.Text, out arr[2]);
                double.TryParse(Thu.Text, out arr[3]);
                double.TryParse(Fri.Text, out arr[4]);
                double.TryParse(Sat.Text, out arr[5]);
                double.TryParse(Sun.Text, out arr[6]);
                Template.HoursPerWeekday = arr;
                Template.StandardStart = TimeSpan.Parse(StandardStartBox.Text);
                Template.LunchBreak = TimeSpan.Parse(LunchBreakBox.Text);
                Template.StandardEnd = TimeSpan.Parse(StandardEndBox.Text);

                // Save template so times/hours persist
                TimeTrackingService.Instance.UpsertTemplate(Template);

                // Diagnostic: enumerate dates and reasons for skipping
                var start = Template.StartDate.Date;
                var end = (Template.EndDate ?? Template.StartDate).Date;
                var sb = new StringBuilder();
                int total = 0; int wouldApply = 0;
                for (var d = start; d <= end; d = d.AddDays(1))
                {
                    total++;
                    int idx = ((int)d.DayOfWeek + 6) % 7; // Monday=0
                    if (Template.HoursPerWeekday != null && Template.HoursPerWeekday.Length == 7 && Template.HoursPerWeekday[idx] == 0)
                    {
                        sb.AppendLine($"{d:yyyy-MM-dd}: skipped (template hours = 0)");
                        continue;
                    }

                    var existing = TimeTrackingService.Instance.GetEvents().Any(ev => ev.Timestamp.Date == d && (ev.Type == TimeEventType.Open || ev.Type == TimeEventType.Close));
                    if (existing)
                    {
                        sb.AppendLine($"{d:yyyy-MM-dd}: will overwrite existing events");
                        wouldApply++;
                        continue;
                    }

                    var openDt = d.Add(Template.StandardStart);
                    var closeDt = d.Add(Template.StandardEnd);
                    if (closeDt <= openDt)
                    {
                        sb.AppendLine($"{d:yyyy-MM-dd}: skipped (invalid start/end times)");
                        continue;
                    }

                    sb.AppendLine($"{d:yyyy-MM-dd}: will apply (open {openDt:HH:mm}, close {closeDt:HH:mm})");
                    wouldApply++;
                }

                // Apply with overwrite so existing events are replaced
                var added = TimeTrackingService.Instance.ApplyTemplate(Template, overwriteExisting: true);

                if (added == 0)
                {
                    sb.Insert(0, $"Applied template to 0 days. Analysis for range {start:yyyy-MM-dd}..{end:yyyy-MM-dd} (total {total}):\n\n");
                    MessageBox.Show(sb.ToString(), "Apply Result", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show($"Applied template to {added} day(s).\n\nDetails:\n" + sb.ToString(), "Apply", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to apply template: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
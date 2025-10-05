using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Input;

namespace Todo
{
    public partial class ShiftDialog : Window
    {
        public Shift Shift { get; private set; }
        private DateTime _forDate;

        private bool _startManual = false;
        private bool _endManual = false;

        public ShiftDialog() : this(null, DateTime.Today)
        {
        }

        public ShiftDialog(Shift s, DateTime forDate)
        {
            InitializeComponent();
            _forDate = forDate.Date;

            // remember whether this was passed in (editing) or created here (new)
            bool isNew = s == null;

            // default lunch 1 hour if creating new
            if (isNew)
                s = new Shift { Start = TimeSpan.FromHours(9), End = TimeSpan.FromHours(17), Description = string.Empty, LunchBreak = TimeSpan.FromHours(1) };

            Shift = s;

            var startList = (ListView)this.FindName("StartList");
            var endList = (ListView)this.FindName("EndList");
            var startText = (TextBox)this.FindName("StartText");
            var endText = (TextBox)this.FindName("EndText");
            var descBox = (TextBox)this.FindName("DescBox");
            var startManualInfo = (TextBlock)this.FindName("StartManualInfo");
            var endManualInfo = (TextBlock)this.FindName("EndManualInfo");
            var lunchText = (TextBox)this.FindName("LunchText");
            var dayModeCombo = (ComboBox)this.FindName("DayModeCombo");

            var opens = TimeTrackingService.Instance.GetEvents().Where(e => e.Type == TimeEventType.Open && e.Timestamp.Date == _forDate && !e.Generated).Select(e => e.Timestamp.TimeOfDay).Distinct().OrderBy(t => t).ToList();
            var closes = TimeTrackingService.Instance.GetEvents().Where(e => e.Type == TimeEventType.Close && e.Timestamp.Date == _forDate && !e.Generated).Select(e => e.Timestamp.TimeOfDay).Distinct().OrderBy(t => t).ToList();

            // wrap times into items
            var startItems = opens.Select(t => new TimeItem { Time = t, Display = t.ToString(@"hh\:mm") }).ToList();
            var endItems = closes.Select(t => new TimeItem { Time = t, Display = t.ToString(@"hh\:mm") }).ToList();

            if (startList != null)
            {
                startList.ItemsSource = startItems;
                var style = new Style(typeof(ListViewItem));
                style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(6)));
                style.Setters.Add(new Setter(Control.BorderBrushProperty, Brushes.LightGray));
                style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
                var selTrigger = new Trigger { Property = ListViewItem.IsSelectedProperty, Value = true };
                selTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, Brushes.DodgerBlue));
                selTrigger.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(2)));
                selTrigger.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.Bold));
                style.Triggers.Add(selTrigger);
                startList.ItemContainerStyle = style;
            }

            if (endList != null)
            {
                endList.ItemsSource = endItems;
                var style = new Style(typeof(ListViewItem));
                style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(6)));
                style.Setters.Add(new Setter(Control.BorderBrushProperty, Brushes.LightGray));
                style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
                var selTrigger = new Trigger { Property = ListViewItem.IsSelectedProperty, Value = true };
                selTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, Brushes.DodgerBlue));
                selTrigger.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(2)));
                selTrigger.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.Bold));
                style.Triggers.Add(selTrigger);
                endList.ItemContainerStyle = style;
            }

            // If editing an existing shift, prefer the shift's saved start/end values and select matching list items if available.
            if (!isNew)
            {
                // start
                var shiftStartDisplay = Shift.Start.ToString(@"hh\:mm");
                bool startMatched = false;
                if (startList != null && startItems.Count > 0)
                {
                    for (int i = 0; i < startItems.Count; i++)
                    {
                        if (startItems[i].Display == shiftStartDisplay)
                        {
                            startList.SelectedIndex = i;
                            startText.Text = startItems[i].Display;
                            startMatched = true;
                            break;
                        }
                    }
                }
                if (!startMatched && startText != null)
                    startText.Text = shiftStartDisplay;

                // end
                var shiftEndDisplay = Shift.End.ToString(@"hh\:mm");
                bool endMatched = false;
                if (endList != null && endItems.Count > 0)
                {
                    for (int i = 0; i < endItems.Count; i++)
                    {
                        if (endItems[i].Display == shiftEndDisplay)
                        {
                            endList.SelectedIndex = i;
                            endText.Text = endItems[i].Display;
                            endMatched = true;
                            break;
                        }
                    }
                }
                if (!endMatched && endText != null)
                    endText.Text = shiftEndDisplay;
            }
            else
            {
                // defaults prefer earliest real open / latest real close when creating new
                if (startItems.Count > 0 && startList != null)
                {
                    startList.SelectedIndex = 0;
                    startText.Text = startItems[0].Display;
                }
                else if (startText != null)
                    startText.Text = Shift.Start.ToString(@"hh\:mm");

                if (endItems.Count > 0 && endList != null)
                {
                    endList.SelectedIndex = endItems.Count - 1;
                    endText.Text = endItems.Last().Display;
                }
                else if (endText != null)
                    endText.Text = Shift.End.ToString(@"hh\:mm");
            }

            if (descBox != null) descBox.Text = Shift.Description;

            _startManual = s != null ? s.ManualStartOverride : false;
            _endManual = s != null ? s.ManualEndOverride : false;

            if (_startManual && startText != null && string.IsNullOrWhiteSpace(startText.Text)) startText.Text = Shift.Start.ToString(@"hh\:mm");
            if (_endManual && endText != null && string.IsNullOrWhiteSpace(endText.Text)) endText.Text = Shift.End.ToString(@"hh\:mm");

            if (startManualInfo != null) startManualInfo.Visibility = _startManual ? Visibility.Visible : Visibility.Collapsed;
            if (endManualInfo != null) endManualInfo.Visibility = _endManual ? Visibility.Visible : Visibility.Collapsed;

            // Set lunch and day mode from Shift values. Always show the saved LunchBreak value (even 00:00).
            if (lunchText != null) lunchText.Text = Shift.LunchBreak.ToString(@"hh\:mm");
            if (dayModeCombo != null && s != null)
            {
                var mode = s.DayMode ?? "Regular Day";
                foreach (ComboBoxItem item in dayModeCombo.Items)
                {
                    if ((item.Content as string) == mode) { item.IsSelected = true; break; }
                }
            }
        }

        private void StartList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _startManual = false;
            var startText = (TextBox)this.FindName("StartText");
            var startManualInfo = (TextBlock)this.FindName("StartManualInfo");
            var list = sender as ListView;
            if (list != null && list.SelectedItem is TimeItem ti && startText != null)
            {
                startText.Text = ti.Display;
            }
            if (startManualInfo != null) startManualInfo.Visibility = Visibility.Collapsed;
        }

        private void EndList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _endManual = false;
            var endText = (TextBox)this.FindName("EndText");
            var endManualInfo = (TextBlock)this.FindName("EndManualInfo");
            var list = sender as ListView;
            if (list != null && list.SelectedItem is TimeItem ti && endText != null)
            {
                endText.Text = ti.Display;
            }
            if (endManualInfo != null) endManualInfo.Visibility = Visibility.Collapsed;
        }

        private void StartText_TextChanged(object sender, TextChangedEventArgs e)
        {
            var startTextBox = (TextBox)this.FindName("StartText");
            var list = (ListView)this.FindName("StartList");
            var startManualInfo = (TextBlock)this.FindName("StartManualInfo");

            // If the text matches an available list item, consider this NOT manual (user chose from list).
            bool matchesList = false;
            if (list != null && startTextBox != null && !string.IsNullOrWhiteSpace(startTextBox.Text))
            {
                foreach (var it in list.Items)
                {
                    if (it is TimeItem ti && ti.Display == startTextBox.Text) { matchesList = true; break; }
                }
            }

            _startManual = !matchesList;
            if (list != null && matchesList == false) list.SelectedIndex = -1; // clear selection when user typed a manual value
            if (startManualInfo != null) startManualInfo.Visibility = _startManual ? Visibility.Visible : Visibility.Collapsed;
            UpdateSelectionHighlightForList("StartList", startTextBox?.Text);
        }

        private void EndText_TextChanged(object sender, TextChangedEventArgs e)
        {
            var endTextBox = (TextBox)this.FindName("EndText");
            var list = (ListView)this.FindName("EndList");
            var endManualInfo = (TextBlock)this.FindName("EndManualInfo");

            bool matchesList = false;
            if (list != null && endTextBox != null && !string.IsNullOrWhiteSpace(endTextBox.Text))
            {
                foreach (var it in list.Items)
                {
                    if (it is TimeItem ti && ti.Display == endTextBox.Text) { matchesList = true; break; }
                }
            }

            _endManual = !matchesList;
            if (list != null && matchesList == false) list.SelectedIndex = -1;
            if (endManualInfo != null) endManualInfo.Visibility = _endManual ? Visibility.Visible : Visibility.Collapsed;
            UpdateSelectionHighlightForList("EndList", endTextBox?.Text);
        }

        private void UpdateSelectionHighlightForList(string listName, string text)
        {
            var list = (ListView)this.FindName(listName);
            if (list == null) return;
            if (string.IsNullOrWhiteSpace(text)) { list.SelectedIndex = -1; return; }
            for (int i = 0; i < list.Items.Count; i++)
            {
                var it = list.Items[i] as TimeItem;
                if (it != null && it.Display == text)
                {
                    if ((listName == "StartList" && !_startManual) || (listName == "EndList" && !_endManual))
                    {
                        list.SelectedIndex = i;
                        return;
                    }
                }
            }
            list.SelectedIndex = -1;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var startText = (TextBox)this.FindName("StartText");
            var endText = (TextBox)this.FindName("EndText");
            var descBox = (TextBox)this.FindName("DescBox");
            var lunchText = (TextBox)this.FindName("LunchText");
            var dayModeCombo = (ComboBox)this.FindName("DayModeCombo");

            var startVal = startText != null ? startText.Text : string.Empty;
            var endVal = endText != null ? endText.Text : string.Empty;

            if (!TimeSpan.TryParse(startVal, out var st)) { MessageBox.Show("Invalid start time. Use HH:MM format.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            if (!TimeSpan.TryParse(endVal, out var et)) { MessageBox.Show("Invalid end time. Use HH:MM format.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            if (et <= st) { MessageBox.Show("End time must be after start time.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning); return; }

            Shift.Start = st;
            Shift.End = et;
            if (descBox != null) Shift.Description = descBox.Text ?? string.Empty;

            // parse lunch
            if (lunchText != null && TimeSpan.TryParse(lunchText.Text, out var lb)) Shift.LunchBreak = lb; else Shift.LunchBreak = TimeSpan.Zero;

            // day mode
            if (dayModeCombo != null && dayModeCombo.SelectedItem is ComboBoxItem cbi) Shift.DayMode = cbi.Content as string;

            Shift.ManualStartOverride = _startManual;
            Shift.ManualEndOverride = _endManual;

            this.DialogResult = true;
            this.Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        // Handle Enter to accept and Escape to cancel anywhere in the window
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                Ok_Click(this, new RoutedEventArgs());
            }
            else if (e.Key == Key.Escape)
            {
                e.Handled = true;
                Cancel_Click(this, new RoutedEventArgs());
            }
        }

        private class TimeItem
        {
            public TimeSpan Time { get; set; }
            public string Display { get; set; }
            public override string ToString() => Display;
        }
    }
}

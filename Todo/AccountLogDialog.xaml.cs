using System;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using System.Text;
using System.Collections.Generic;

namespace Todo
{
    public partial class AccountLogDialog : Window
    {
        private List<AccountLogEntry> _currentLogItems = new List<AccountLogEntry>();

        public AccountLogDialog(DateTime? from = null, DateTime? to = null, string kind = null)
        {
            InitializeComponent();
            FromDatePicker.SelectedDate = from ?? DateTime.Today.AddYears(-1);
            ToDatePicker.SelectedDate = to ?? DateTime.Today;

            KindCombo.Items.Add("(All)");
            KindCombo.Items.Add("Holiday");
            KindCombo.Items.Add("TIL");
            KindCombo.SelectedItem = string.IsNullOrEmpty(kind) ? "(All)" : kind;

            LoadItems();
        }

        private class DisplayAccountLogEntry
        {
            public DateTime Date { get; set; }
            public string Kind { get; set; }
            public string DeltaDisplay { get; set; }
            public string BalanceDisplay { get; set; }
            public string Note { get; set; }
            public string AffectedDateDisplay { get; set; }
        }

        private static bool IsManualSet(AccountLogEntry a)
        {
            var n = a?.Note ?? string.Empty;
            return n.IndexOf("manual set", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void LoadItems()
        {
            var f = FromDatePicker.SelectedDate ?? DateTime.Today.AddYears(-1);
            var t = ToDatePicker.SelectedDate ?? DateTime.Today;
            string k = (KindCombo.SelectedItem as string);
            if (k == "(All)") k = null;

            // Get saved log entries (respecting kind filter)
            var saved = TimeTrackingService.Instance.GetAccountLogEntries(f, t, k).OrderBy(a => a.Date).ToList();

            // Compute running balances for both TIL and Holiday entries (saved only)
            try
            {
                var asc = saved.OrderBy(a => a.Date).ThenBy(a => a.Kind).ToList();

                double startTIL = ComputeStartRunning("TIL", f);
                double startHol = ComputeStartRunning("Holiday", f);
                double runningTIL = 0.0; double runningHol = 0.0;

                foreach (var entry in asc)
                {
                    if (string.Equals(entry.Kind, "TIL", StringComparison.OrdinalIgnoreCase))
                    {
                        if (IsManualSet(entry)) { runningTIL = 0.0; entry.Balance = startTIL + runningTIL; }
                        else { runningTIL += entry.Delta; entry.Balance = startTIL + runningTIL; }
                    }
                    else if (string.Equals(entry.Kind, "Holiday", StringComparison.OrdinalIgnoreCase))
                    {
                        if (IsManualSet(entry)) { runningHol = 0.0; entry.Balance = startHol + runningHol; }
                        else { runningHol += entry.Delta; entry.Balance = startHol + runningHol; }
                    }
                }

                _currentLogItems = asc.OrderByDescending(a => a.Date).ThenByDescending(a => a.Kind).ToList();
            }
            catch
            {
                _currentLogItems = saved.OrderByDescending(a => a.Date).ThenByDescending(a => a.Kind).ToList();
            }

            // Map to display entries with units for TIL/Holiday
            var display = _currentLogItems.Select(a => new DisplayAccountLogEntry
            {
                Date = a.Date,
                Kind = a.Kind,
                DeltaDisplay = FormatDelta(a),
                BalanceDisplay = FormatBalance(a),
                Note = a.Note,
                AffectedDateDisplay = a.AffectedDate.HasValue ? a.AffectedDate.Value.ToString("yyyy-MM-dd") : string.Empty
            }).ToList();

            LogList.ItemsSource = display;
        }

        private static double ComputeStartRunning(string kind, DateTime startDate)
        {
            double running = 0.0;
            try
            {
                var allBefore = TimeTrackingService.Instance.GetAccountLogEntries(DateTime.MinValue, startDate.AddDays(-1), kind).OrderBy(a => a.Date).ToList();
                var lastSet = allBefore.Where(IsManualSet).OrderBy(a => a.Date).LastOrDefault();
                if (lastSet != null)
                {
                    running = lastSet.Balance;
                    foreach (var e in allBefore.Where(x => x.Date > lastSet.Date && !IsManualSet(x))) running += e.Delta;
                }
                else
                {
                    var acc = TimeTrackingService.Instance.GetAccountState();
                    running = string.Equals(kind, "TIL", StringComparison.OrdinalIgnoreCase) ? acc.TILOffset : acc.HolidayOffset;
                    foreach (var e in allBefore.Where(x => !IsManualSet(x))) running += e.Delta;
                }
            }
            catch { }
            return running;
        }

        private string FormatDelta(AccountLogEntry a)
        {
            if (a == null) return string.Empty;
            if (string.Equals(a.Kind, "TIL", StringComparison.OrdinalIgnoreCase))
                return a.Delta.ToString("0.##") + "h";
            if (string.Equals(a.Kind, "Holiday", StringComparison.OrdinalIgnoreCase))
                return a.Delta.ToString("0.##") + "d";
            return a.Delta.ToString("0.##");
        }

        private string FormatBalance(AccountLogEntry a)
        {
            if (a == null) return string.Empty;
            if (double.IsNaN(a.Balance)) return string.Empty;
            if (string.Equals(a.Kind, "TIL", StringComparison.OrdinalIgnoreCase))
                return a.Balance.ToString("0.##") + "h";
            if (string.Equals(a.Kind, "Holiday", StringComparison.OrdinalIgnoreCase))
                return a.Balance.ToString("0.##") + "d";
            return a.Balance.ToString("0.##");
        }

        private void ApplyFilter_Click(object sender, RoutedEventArgs e)
        {
            LoadItems();
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new SaveFileDialog { Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*", FileName = "account_log.csv" };
                if (dlg.ShowDialog(this) == true)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("Date,Kind,Delta,Balance,Note,AffectedDate");
                    foreach (var a in _currentLogItems)
                    {
                        var date = a.Date.ToString("yyyy-MM-dd HH:mm");
                        var note = (a.Note ?? string.Empty).Replace("\"", "\"\"");
                        var affected = a.AffectedDate.HasValue ? a.AffectedDate.Value.ToString("yyyy-MM-dd") : string.Empty;
                        var deltaStr = FormatDelta(a);
                        var balanceStr = FormatBalance(a);
                        sb.AppendLine($"\"{date}\",{a.Kind},{deltaStr},{balanceStr},\"{note}\",{affected}");
                    }
                    System.IO.File.WriteAllText(dlg.FileName, sb.ToString(), System.Text.Encoding.UTF8);
                    MessageBox.Show("Exported.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to export: " + ex.Message, "Export", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}

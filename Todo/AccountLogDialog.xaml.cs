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

        private void LoadItems()
        {
            var f = FromDatePicker.SelectedDate ?? DateTime.Today.AddYears(-1);
            var t = ToDatePicker.SelectedDate ?? DateTime.Today;
            string k = (KindCombo.SelectedItem as string);
            if (k == "(All)") k = null;

            // Get saved log entries (respecting kind filter)
            var saved = TimeTrackingService.Instance.GetAccountLogEntries(f, t, k).ToList();

            var combined = new List<AccountLogEntry>();
            combined.AddRange(saved);

            // If showing all kinds or specifically TIL, include generated per-day TIL deltas so users can see positive and negative daily changes
            if (string.IsNullOrEmpty(k) || string.Equals(k, "TIL", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var daySummaries = TimeTrackingService.Instance.GetDaySummaries(f, t).OrderBy(d => d.Date).ToList();
                    foreach (var ds in daySummaries)
                    {
                        if (!ds.DeltaHours.HasValue) continue;
                        var delta = ds.DeltaHours.Value;
                        if (Math.Abs(delta) < 0.0001) continue; // skip zero deltas

                        // create a generated account log entry for display/export purposes
                        var gen = new AccountLogEntry
                        {
                            Date = ds.Date, // use the day date as the entry date
                            Kind = "TIL",
                            Delta = delta,
                            Balance = double.NaN, // will be filled later when computing running balance
                            Note = "Daily delta (computed)",
                            AffectedDate = ds.Date
                        };
                        combined.Add(gen);
                    }
                }
                catch { }
            }

            // We will compute running balances for TIL entries (saved + generated)
            try
            {
                // get all saved TIL entries (unfiltered) to determine previous balance state
                var allSavedTIL = TimeTrackingService.Instance.GetAccountLogEntries(DateTime.MinValue, DateTime.MaxValue, "TIL").OrderBy(a => a.Date).ToList();
                // find last saved TIL balance strictly before the display start
                var lastBefore = allSavedTIL.Where(a => a.Date.Date < f.Date).OrderByDescending(a => a.Date).FirstOrDefault();
                double startBalance = 0.0;
                if (lastBefore != null)
                {
                    startBalance = lastBefore.Balance;
                }
                else
                {
                    // If no prior saved balance, attempt to use earliest saved TIL balance as a reference and roll back, otherwise use 0
                    var earliestSaved = allSavedTIL.FirstOrDefault();
                    if (earliestSaved != null)
                    {
                        // compute cumulative from earliestSaved up to entries we care about by treating earliestSaved as base
                        // but for simplicity use 0 base when no prior saved balance exists
                        startBalance = 0.0;
                    }
                    else
                    {
                        startBalance = 0.0;
                    }
                }

                // Build combined list ordered ascending for running calculation
                var asc = combined.OrderBy(a => a.Date).ThenBy(a => a.Kind).ToList();
                double running = 0.0; // running delta since startBalance
                // If there is a saved 'lastBefore' we should incorporate any saved deltas between lastBefore and f, but since lastBefore is before f, running starts at 0
                // Now process entries in chronological order and set Balance for TIL entries
                foreach (var entry in asc)
                {
                    if (string.Equals(entry.Kind, "TIL", StringComparison.OrdinalIgnoreCase))
                    {
                        running += entry.Delta;
                        entry.Balance = startBalance + running;
                    }
                    // leave Holiday balances as-is (they already store Balance)
                }

                // After computing, set the _currentLogItems to combined sorted descending for display
                _currentLogItems = asc.OrderByDescending(a => a.Date).ThenByDescending(a => a.Kind).ToList();
            }
            catch
            {
                // Fallback: show combined unsorted if computation fails
                _currentLogItems = combined.OrderByDescending(a => a.Date).ThenByDescending(a => a.Kind).ToList();
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
            // If balance is NaN we treat it as unknown (shouldn't usually happen after computation)
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
                        // Format delta/balance with units for export (TIL -> hours 'h', Holiday -> days 'd')
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

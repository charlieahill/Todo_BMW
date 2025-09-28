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
            var items = TimeTrackingService.Instance.GetAccountLogEntries(f, t, k);
            _currentLogItems = items.ToList();

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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace Todo
{
    public partial class YearSelectionDialog : Window
    {
        public class YearItem
        {
            public int Year { get; set; }
            public bool IsChecked { get; set; }
        }

        private List<YearItem> _years = new List<YearItem>();
        public IReadOnlyList<int> SelectedYears => _years.Where(y => y.IsChecked).Select(y => y.Year).OrderBy(y => y).ToList();

        public YearSelectionDialog(IEnumerable<int> years, IEnumerable<int> prechecked = null)
        {
            InitializeComponent();
            var set = new HashSet<int>(prechecked ?? Enumerable.Empty<int>());
            _years = years.Distinct().OrderBy(y => y).Select(y => new YearItem { Year = y, IsChecked = set.Contains(y) }).ToList();
            YearsList.ItemsSource = _years;
            UpdateSelectAllState();
        }

        private void UpdateSelectAllState()
        {
            if (_years.Count == 0) { SelectAllYears.IsChecked = false; return; }
            int checkedCount = _years.Count(y => y.IsChecked);
            if (checkedCount == 0) SelectAllYears.IsChecked = false;
            else if (checkedCount == _years.Count) SelectAllYears.IsChecked = true;
            else SelectAllYears.IsChecked = null; // indeterminate
        }

        private void SelectAllYears_Click(object sender, RoutedEventArgs e)
        {
            bool? target = SelectAllYears.IsChecked;
            bool value = target == true;
            foreach (var y in _years) y.IsChecked = value;
            YearsList.Items.Refresh();
            UpdateSelectAllState();
        }

        private void YearItem_CheckChanged(object sender, RoutedEventArgs e)
        {
            UpdateSelectAllState();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!SelectedYears.Any())
            {
                MessageBox.Show("Select at least one year.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            this.DialogResult = true;
        }
    }
}

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
        }

        private void CheckAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var y in _years) y.IsChecked = true;
            YearsList.Items.Refresh();
        }

        private void UncheckAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var y in _years) y.IsChecked = false;
            YearsList.Items.Refresh();
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

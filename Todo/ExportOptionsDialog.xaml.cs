using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Todo
{
    public partial class ExportOptionsDialog : Window
    {
        public class YearItem { public int Year { get; set; } public bool IsChecked { get; set; } }
        public class MapItem { public string Name { get; set; } public bool IsChecked { get; set; } public string Key { get; set; } }

        private List<YearItem> _years = new();
        private List<MapItem> _maps = new();

        public IReadOnlyList<int> SelectedYears => _years.Where(y => y.IsChecked).Select(y => y.Year).OrderBy(y => y).ToList();
        public bool IncludeDayType => _maps.Any(m => m.Key == "day" && m.IsChecked);
        public bool IncludeLocation => _maps.Any(m => m.Key == "loc" && m.IsChecked);
        public bool IncludeOvertime => _maps.Any(m => m.Key == "ot" && m.IsChecked);

        public ExportOptionsDialog(IEnumerable<int> years, IEnumerable<int> precheckedYears = null, bool defaultDayType = false, bool defaultLocation = true, bool defaultOvertime = true)
        {
            InitializeComponent();

            var set = new HashSet<int>(precheckedYears ?? Array.Empty<int>());
            _years = years.Distinct().OrderBy(y => y).Select(y => new YearItem { Year = y, IsChecked = set.Contains(y) }).ToList();
            YearsList.ItemsSource = _years;
            UpdateSelectAllYearsState();

            _maps = new List<MapItem>
            {
                new MapItem { Name = "Day type map", Key = "day", IsChecked = defaultDayType },
                new MapItem { Name = "Physical location map", Key = "loc", IsChecked = defaultLocation },
                new MapItem { Name = "Overtime gradient map", Key = "ot", IsChecked = defaultOvertime },
            };
            MapsList.ItemsSource = _maps;
            UpdateSelectAllMapsState();
        }

        private void UpdateSelectAllYearsState()
        {
            if (_years.Count == 0) { SelectAllYears.IsChecked = false; return; }
            int c = _years.Count(y => y.IsChecked);
            if (c == 0) SelectAllYears.IsChecked = false;
            else if (c == _years.Count) SelectAllYears.IsChecked = true;
            else SelectAllYears.IsChecked = null;
        }

        private void UpdateSelectAllMapsState()
        {
            if (_maps.Count == 0) { SelectAllMaps.IsChecked = false; return; }
            int c = _maps.Count(m => m.IsChecked);
            if (c == 0) SelectAllMaps.IsChecked = false;
            else if (c == _maps.Count) SelectAllMaps.IsChecked = true;
            else SelectAllMaps.IsChecked = null;
        }

        private void SelectAllYears_Click(object sender, RoutedEventArgs e)
        {
            bool value = SelectAllYears.IsChecked == true;
            foreach (var y in _years) y.IsChecked = value;
            YearsList.Items.Refresh();
            UpdateSelectAllYearsState();
        }

        private void SelectAllMaps_Click(object sender, RoutedEventArgs e)
        {
            bool value = SelectAllMaps.IsChecked == true;
            foreach (var m in _maps) m.IsChecked = value;
            MapsList.Items.Refresh();
            UpdateSelectAllMapsState();
        }

        private void YearItem_CheckChanged(object sender, RoutedEventArgs e) => UpdateSelectAllYearsState();
        private void MapItem_CheckChanged(object sender, RoutedEventArgs e) => UpdateSelectAllMapsState();

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!SelectedYears.Any())
            {
                MessageBox.Show("Select at least one year.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (!_maps.Any(m => m.IsChecked))
            {
                MessageBox.Show("Select at least one map type.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            this.DialogResult = true;
        }
    }
}

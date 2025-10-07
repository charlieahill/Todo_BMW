using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Todo
{
    public partial class LocationColorsDialog : Window
    {
        public class LocationColorItem
        {
            public string Location { get; set; }
            public string ColorHex { get; set; }
            public Brush ColorBrush
            {
                get
                {
                    try
                    {
                        if (string.IsNullOrWhiteSpace(ColorHex)) return Brushes.White;
                        var c = (Color)ColorConverter.ConvertFromString(ColorHex);
                        return new SolidColorBrush(c);
                    }
                    catch { return Brushes.White; }
                }
            }
        }

        private List<LocationColorItem> _items = new List<LocationColorItem>();

        public LocationColorsDialog()
        {
            InitializeComponent();
            LoadItems();
        }

        private void LoadItems()
        {
            try
            {
                var allLocs = TimeTrackingService.Instance.DiscoverPhysicalLocations();
                var map = TimeTrackingService.Instance.GetLocationColors();
                _items = allLocs.Select(loc => new LocationColorItem
                {
                    Location = loc,
                    ColorHex = map.TryGetValue(loc, out var hex) ? hex : string.Empty
                }).ToList();
                ColorsGrid.ItemsSource = _items;
            }
            catch { }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (var it in _items)
                {
                    if (string.IsNullOrWhiteSpace(it.Location)) continue;
                    if (string.IsNullOrWhiteSpace(it.ColorHex))
                        TimeTrackingService.Instance.RemoveLocationColor(it.Location);
                    else
                        TimeTrackingService.Instance.SetLocationColor(it.Location, it.ColorHex);
                }
                this.DialogResult = true;
            }
            catch
            {
                this.DialogResult = true;
            }
        }

        private void PickColor_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button b && b.Tag is LocationColorItem it)
                {
                    // Use WPF ColorDialog via WinForms interop for simplicity
                    var cd = new System.Windows.Forms.ColorDialog();
                    if (!string.IsNullOrWhiteSpace(it.ColorHex))
                    {
                        try
                        {
                            var c = (Color)ColorConverter.ConvertFromString(it.ColorHex);
                            cd.Color = System.Drawing.Color.FromArgb(c.A, c.R, c.G, c.B);
                        }
                        catch { }
                    }
                    if (cd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        var c = cd.Color;
                        it.ColorHex = string.Format("#{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B);
                        ColorsGrid.Items.Refresh();
                    }
                }
            }
            catch { }
        }

        private void ClearColor_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button b && b.Tag is LocationColorItem it)
                {
                    it.ColorHex = string.Empty;
                    ColorsGrid.Items.Refresh();
                }
            }
            catch { }
        }
    }
}

using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Data;

namespace Todo
{
    public class LinkTypeToVisibilityConverter : IValueConverter
    {
        // converterParameter expected: "file", "folder", or "web"
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                var param = (parameter ?? string.Empty).ToString().ToLowerInvariant();
                var path = (value as string) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(path)) return Visibility.Collapsed;

                switch (param)
                {
                    case "folder":
                        try { return Directory.Exists(path) ? Visibility.Visible : Visibility.Collapsed; } catch { return Visibility.Collapsed; }
                    case "file":
                        try { return File.Exists(path) ? Visibility.Visible : Visibility.Collapsed; } catch { return Visibility.Collapsed; }
                    case "web":
                        try
                        {
                            if (Uri.TryCreate(path, UriKind.Absolute, out var u))
                            {
                                var s = u.Scheme?.ToLowerInvariant();
                                if (s == "http" || s == "https") return Visibility.Visible;
                            }
                            return Visibility.Collapsed;
                        }
                        catch { return Visibility.Collapsed; }
                    default:
                        return Visibility.Collapsed;
                }
            }
            catch
            {
                return Visibility.Collapsed;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Todo
{
    public class CountToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string param = (parameter ?? string.Empty).ToString().ToLowerInvariant();

            if (value == null)
                return Visibility.Collapsed;

            // If value is string (Description)
            if (value is string s)
            {
                if (param == "desc")
                    return !string.IsNullOrWhiteSpace(s) ? Visibility.Visible : Visibility.Collapsed;
                // fallback
                return !string.IsNullOrWhiteSpace(s) ? Visibility.Visible : Visibility.Collapsed;
            }

            // If value is int (Count)
            if (value is int count)
            {
                switch (param)
                {
                    case "people":
                        return count == 1 ? Visibility.Visible : Visibility.Collapsed;
                    case "group":
                        return count > 1 ? Visibility.Visible : Visibility.Collapsed;
                    case "meet":
                    case "meeting":
                        return count > 0 ? Visibility.Visible : Visibility.Collapsed;
                    default:
                        return count > 0 ? Visibility.Visible : Visibility.Collapsed;
                }
            }

            // If value is a collection (ICollection) - try to get Count property via reflection
            try
            {
                var prop = value.GetType().GetProperty("Count");
                if (prop != null)
                {
                    var cntObj = prop.GetValue(value);
                    if (cntObj is int cnt)
                    {
                        switch (param)
                        {
                            case "people":
                                return cnt == 1 ? Visibility.Visible : Visibility.Collapsed;
                            case "group":
                                return cnt > 1 ? Visibility.Visible : Visibility.Collapsed;
                            case "meet":
                            case "meeting":
                                return cnt > 0 ? Visibility.Visible : Visibility.Collapsed;
                            default:
                                return cnt > 0 ? Visibility.Visible : Visibility.Collapsed;
                        }
                    }
                }
            }
            catch { }

            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

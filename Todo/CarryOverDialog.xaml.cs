using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;

namespace Todo
{
    public class CarryOverItem : INotifyPropertyChanged
    {
        public TaskModel Task { get; }

        private bool _isCopyToday = true;
        public bool IsCopyToday { get => _isCopyToday; set { if (_isCopyToday == value) return; _isCopyToday = value; OnPropertyChanged(nameof(IsCopyToday)); if (value) { IsCopyFuture = false; IsMarkCompleted = false; } } }

        private bool _isCopyFuture = false;
        public bool IsCopyFuture { get => _isCopyFuture; set { if (_isCopyFuture == value) return; _isCopyFuture = value; OnPropertyChanged(nameof(IsCopyFuture)); if (value) { IsCopyToday = false; IsMarkCompleted = false; } } }

        private bool _isMarkCompleted = false;
        public bool IsMarkCompleted { get => _isMarkCompleted; set { if (_isMarkCompleted == value) return; _isMarkCompleted = value; OnPropertyChanged(nameof(IsMarkCompleted)); if (value) { IsCopyToday = false; IsCopyFuture = false; } } }

        private DateTime? _futureDate = DateTime.Today.AddDays(1);
        public DateTime? FutureDate { get => _futureDate; set { if (_futureDate == value) return; _futureDate = value; OnPropertyChanged(nameof(FutureDate)); } }

        public CarryOverItem(TaskModel task)
        {
            Task = task;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public partial class CarryOverDialog : Window
    {
        public ObservableCollection<CarryOverItem> Items { get; } = new ObservableCollection<CarryOverItem>();

        public CarryOverDialog() { InitializeComponent(); DataContext = this; }

        public CarryOverDialog(System.Collections.Generic.IEnumerable<TaskModel> tasks) : this()
        {
            foreach (var t in tasks) Items.Add(new CarryOverItem(t));
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

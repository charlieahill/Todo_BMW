using System.Windows;

namespace Todo
{
    public partial class SettingsDialog : Window
    {
        public SettingsDialog()
        {
            InitializeComponent();
            LoadFromSettings();
        }

        private void LoadFromSettings()
        {
            var s = SettingsService.Instance.Settings;
            MoveToTopRadio.IsChecked = s.MoveBehavior == MoveCompletedBehavior.MoveToTop;
            DoNotMoveRadio.IsChecked = s.MoveBehavior == MoveCompletedBehavior.DoNotMove;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var s = SettingsService.Instance.Settings;
            if (MoveToTopRadio.IsChecked == true)
                s.MoveBehavior = MoveCompletedBehavior.MoveToTop;
            else if (DoNotMoveRadio.IsChecked == true)
                s.MoveBehavior = MoveCompletedBehavior.DoNotMove;

            SettingsService.Instance.Save();
            DialogResult = true;
            Close();
        }
    }
}

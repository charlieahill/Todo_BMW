using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Globalization;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

namespace Todo
{
    public partial class TimeTrackingDialog : Window
    {
        private List<DayViewModel> _currentDays = new List<DayViewModel>();
        // track the currently active/selected day so we can persist edits when selection changes
        private DayViewModel _activeDay = null;

        // Toggle to colour month grid by physical location instead of day type
        public static bool ColorByPhysicalLocation { get; set; } = false;

        // Temporary visual debug: set to true to highlight and show diagnostic info for hours worked
        // private const bool DebugShowTotals = true;

        private const string WindowSettingsFile = "timetracking_window.json";

        // Helper colors map for DayType palette (same as DayViewModel.BackgroundBrush)
        private static readonly Dictionary<DayType, XColor> DayTypePdfColors = new Dictionary<DayType, XColor>
        {
            { DayType.WorkingDay, XColor.FromArgb(0xFF, 0xDF, 0xF0, 0xD8) },
            { DayType.Weekend, XColor.FromArgb(0xFF, 0xF0, 0xF0, 0xF0) },
            { DayType.PublicHoliday, XColor.FromArgb(0xFF, 0xFF, 0xE5, 0xE5) },
            { DayType.Vacation, XColor.FromArgb(0xFF, 0xDD, 0xEE, 0xFF) },
            { DayType.TimeInLieu, XColor.FromArgb(0xFF, 0xFF, 0xF2, 0xCC) },
            { DayType.Other, XColor.FromArgb(0xFF, 0xEF, 0xEF, 0xEF) }
        };

        public TimeTrackingDialog()
        {
            InitializeComponent();

            // Try to restore previous window position/size
            try
            {
                if (File.Exists(WindowSettingsFile))
                {
                    var j = File.ReadAllText(WindowSettingsFile);
                    var s = JsonSerializer.Deserialize<WindowSettings>(j);
                    if (s != null)
                    {
                        this.WindowStartupLocation = WindowStartupLocation.Manual;
                        this.Left = s.Left;
                        this.Top = s.Top;
                        this.Width = s.Width;
                        this.Height = s.Height;
                        // set state after bounds
                        if (s.State == WindowState.Maximized)
                            this.WindowState = WindowState.Maximized;
                    }
                }
            }
            catch { }

            // save settings when closing
            this.Closing += TimeTrackingDialog_Closing;

            PopulateYearMonth();
            YearCombo.SelectedItem = DateTime.Today.Year.ToString();
            MonthCombo.SelectedIndex = DateTime.Today.Month - 1;
            LoadTemplates();
            SetupDayTypeCombo();
            PopulateDaysForCurrentMonth();
            UpdateSelectedDateLabel(DateTime.Today);

            // populate combobox item lists from templates/overrides
            UpdateComboLists();
        }

        private void TimeTrackingDialog_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                // persist any pending edits
                SaveCurrentDayEdits();

                // If maximized, use RestoreBounds for the saved location/size
                double left = this.Left;
                double top = this.Top;
                double width = this.Width;
                double height = this.Height;
                var state = this.WindowState;
                if (this.WindowState == WindowState.Maximized)
                {
                    var rb = this.RestoreBounds;
                    left = rb.Left;
                    top = rb.Top;
                    width = rb.Width;
                    height = rb.Height;
                }

                var s = new WindowSettings { Left = left, Top = top, Width = width, Height = height, State = state };
                var j = JsonSerializer.Serialize(s, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(WindowSettingsFile, j);
            }
            catch { }
        }

        private void SetupDayTypeCombo()
        {
            DayTypeCombo.ItemsSource = Enum.GetValues(typeof(DayType)).Cast<DayType>().Select(d => d.ToString());
        }

        private void PopulateYearMonth()
        {
            YearCombo.Items.Clear();
            var start = 2020;
            var end = DateTime.Today.Year + 1;
            for (int y = start; y <= end; y++)
                YearCombo.Items.Add(y.ToString());

            MonthCombo.Items.Clear();
            for (int m = 1; m <= 12; m++)
                MonthCombo.Items.Add(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m));
        }

        private void LoadTemplates()
        {
            // Previously populated a ListBox of templates. Templates UI removed; keep templates loaded into combo lists only.
            // Refresh combo lists because templates can provide known positions/locations
            UpdateComboLists();
        }

        // Re-add missing handlers referenced by XAML
        // Removed TemplatesList MouseDoubleClick/Edit/Delete handlers because templates list was removed from UI.

        private void UpdateComboLists()
        {
            try
            {
                var templates = TimeTrackingService.Instance.GetTemplates();
                var overrides = TimeTrackingService.Instance.GetOverrides();

                var positions = new List<string>();
                positions.AddRange(templates.Select(t => t.EmploymentPosition).Where(s => !string.IsNullOrWhiteSpace(s)));
                positions.AddRange(overrides.Select(o => o.Position).Where(s => !string.IsNullOrWhiteSpace(s)));
                PositionCombo.ItemsSource = positions.Distinct().ToList();

                var locations = new List<string>();
                locations.AddRange(templates.Select(t => t.EmploymentLocation).Where(s => !string.IsNullOrWhiteSpace(s)));
                locations.AddRange(overrides.Select(o => o.Location).Where(s => !string.IsNullOrWhiteSpace(s)));
                LocationCombo.ItemsSource = locations.Distinct().ToList();

                var phys = overrides.Select(o => o.PhysicalLocation).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                PhysicalLocationCombo.ItemsSource = phys;
            }
            catch { }
        }

        private void PopulateDaysForCurrentMonth()
        {
            if (YearCombo.SelectedItem == null || MonthCombo.SelectedIndex < 0)
                return;

            int year = int.Parse((string)YearCombo.SelectedItem);
            int month = MonthCombo.SelectedIndex + 1;
            var first = new DateTime(year, month, 1);
            var last = first.AddMonths(1).AddDays(-1);
            var summaries = TimeTrackingService.Instance.GetDaySummaries(first, last);
            _currentDays = summaries.Select(s => new DayViewModel(s)).ToList();

            // Load persisted shifts and then compute cumulative accounts per day (TIL hours and Holiday days)
            // We'll compute running totals in chronological order
            double runningTIL = TimeTrackingService.Instance.GetAccountState().TILOffset;
            double runningHoliday = TimeTrackingService.Instance.GetAccountState().HolidayOffset; // represent days (can be negative)

            // Load persisted overrides for each day
            foreach (var d in _currentDays)
            {
                var ov = TimeTrackingService.Instance.GetOverrideForDate(d.Date);
                if (ov != null)
                {
                    d.PositionOverride = ov.Position;
                    d.LocationOverride = ov.Location;
                    // If override contains a specific physical location, prefer that. Otherwise default to employment/location value
                    if (!string.IsNullOrWhiteSpace(ov.PhysicalLocation))
                        d.PhysicalLocationOverride = ov.PhysicalLocation;
                    else if (!string.IsNullOrWhiteSpace(ov.Location))
                        d.PhysicalLocationOverride = ov.Location;
                    else if (d.Template != null && !string.IsNullOrWhiteSpace(d.Template.Location))
                        d.PhysicalLocationOverride = d.Template.Location;

                    d.TargetHours = ov.TargetHours;
                    // load persisted DayType override if present
                    if (ov.DayType.HasValue)
                    {
                        d.SetDayType(ov.DayType.Value);
                    }
                }
                else
                {
                    // default physical location to template employment location if available
                    if (d.Template != null && !string.IsNullOrWhiteSpace(d.Template.Location))
                        d.PhysicalLocationOverride = d.Template.Location;
                }

                // load persisted shifts for day
                var saved = TimeTrackingService.Instance.GetShiftsForDate(d.Date).ToList();
                if (saved != null && saved.Count > 0)
                {
                    // clone into DayViewModel.Shifts
                    d.Shifts = saved.Select(s => new Shift
                    {
                        Start = s.Start,
                        End = s.End,
                        Description = s.Description,
                        ManualStartOverride = s.ManualStartOverride,
                        ManualEndOverride = s.ManualEndOverride,
                        LunchBreak = s.LunchBreak,
                        DayMode = s.DayMode,
                        Date = s.Date
                    }).ToList();
                }
                // compute worked hours for this day (prefer shifts sum if available)
                double dayWorked = 0;
                if (d.Shifts != null && d.Shifts.Count > 0)
                    dayWorked = d.Shifts.Sum(s => s.Hours);
                else if (d.Worked.HasValue)
                    dayWorked = d.Worked.Value;

                // determine target hours for this day (override or template default)
                double dayTarget = d.TargetHours ?? 0;
                if (dayTarget <= 0 && d.Template != null)
                {
                    int idx = ((int)d.Date.DayOfWeek + 6) % 7;
                    if (d.Template.HoursPerWeekday != null && d.Template.HoursPerWeekday.Length == 7)
                        dayTarget = d.Template.HoursPerWeekday[idx];
                }

                // compute delta = worked - target (can be negative)
                double dayDelta = dayWorked - (dayTarget > 0 ? dayTarget : 0);

                // Update running account values depending on DayType
                if (d.DayType == DayType.WorkingDay)
                {
                    runningTIL += dayDelta; // accumulate delta
                }
                else if (d.DayType == DayType.Vacation)
                {
                    runningHoliday -= 1.0;
                }

                d.CumulativeTIL = runningTIL;
                d.CumulativeHoliday = runningHoliday;
            }

            DaysList.ItemsSource = _currentDays;
            // populate month grid: include leading/trailing blanks to fill weeks
            var items = new List<DayViewModel?>();
            int leading = ((int)first.DayOfWeek + 6) % 7; // Monday=0 approach isn't used here; use Sun=0..Sat=6
            // adjust so that grid starts with Monday (optional) - here start with Monday: calculate days to skip where Monday=0
            // Let's start week on Monday to match list weekday formatting
            leading = (int)first.DayOfWeek - 1; if (leading < 0) leading += 7;
            for (int i = 0; i < leading; i++) items.Add(null);
            foreach (var d in _currentDays) items.Add(d);
            while (items.Count % 7 != 0) items.Add(null);
            MonthGrid.ItemsSource = items;

            // set IsToday flag and reset IsSelected
            var today = DateTime.Today;
            foreach (var d in _currentDays)
            {
                d.IsToday = d.Date.Date == today;
                d.IsSelected = false;
            }

            // select today's date if in range
            var sel = _currentDays.FirstOrDefault(d => d.Date.Date == today);
            if (sel != null)
            {
                DaysList.SelectedItem = sel;
                sel.IsSelected = true;
                _activeDay = sel;
                // update shift totals display for initial selection
                UpdateShiftTotalsDisplay();
            }

            MonthGrid.Items.Refresh();

            // ensure combo lists reflect loaded overrides/templates
            UpdateComboLists();
        }

        private void YearCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (YearCombo.SelectedItem == null) return;
            PopulateDaysForCurrentMonth();
        }

        // Handler referenced from XAML: update DayType when selection changes
        private void DayTypeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (DaysList.SelectedItem is DayViewModel vm && DayTypeCombo.SelectedItem is string sel)
                {
                    if (Enum.TryParse<DayType>(sel, out var dt))
                    {
                        // remember previous type so we can adjust holiday account appropriately
                        var previous = vm.DayType;
                        vm.SetDayType(dt);

                        // Persist the day type change (and other overrides) so it survives restarts
                        try
                        {
                            TimeTrackingService.Instance.UpsertOverride(vm.Date, vm.PositionOverride, vm.LocationOverride, vm.PhysicalLocationOverride, vm.TargetHours, dt);
                        }
                        catch { }

                        // If day marked as holiday, set target hours to zero and persist override
                        if (dt == DayType.Vacation)
                        {
                            vm.TargetHours = 0;
                            try { TimeTrackingService.Instance.UpsertOverride(vm.Date, vm.PositionOverride, vm.LocationOverride, vm.PhysicalLocationOverride, vm.TargetHours, dt); } catch { }
                        }

                        // Adjust global holiday account only when transitioning into/out-of Vacation to avoid double-counting
                        double holidayDelta = 0;
                        if (previous != DayType.Vacation && dt == DayType.Vacation)
                            holidayDelta = -1.0; // taking a vacation day reduces remaining holidays
                        else if (previous == DayType.Vacation && dt != DayType.Vacation)
                            holidayDelta = 1.0; // reverting a vacation restores a day

                        if (holidayDelta != 0)
                        {
                            try
                            {
                                var acc = TimeTrackingService.Instance.GetAccountState();
                                double old = acc.HolidayOffset;
                                acc.HolidayOffset = old + holidayDelta;
                                TimeTrackingService.Instance.SetAccountState(acc);
                                var entry = new AccountLogEntry
                                {
                                    Date = DateTime.Now,
                                    Kind = "Holiday",
                                    Delta = holidayDelta,
                                    Balance = acc.HolidayOffset,
                                    Note = holidayDelta < 0 ? "Vacation taken" : "Vacation restored",
                                    AffectedDate = vm.Date
                                };
                                TimeTrackingService.Instance.AddAccountLogEntry(entry);
                            }
                            catch { }
                        }

                        DaysList.Items.Refresh();
                        MonthGrid.Items.Refresh();
                        // Recompute cumulatives and update accounts since day types/targets changed
                        RecomputeCumulatives();
                        UpdateAccountsDisplay();
                    }
                }
            }
            catch { }
        }

        private void MonthCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MonthCombo.SelectedIndex < 0) return;
            PopulateDaysForCurrentMonth();
        }

        private void TodayButton_Click(object sender, RoutedEventArgs e)
        {
            YearCombo.SelectedItem = DateTime.Today.Year.ToString();
            MonthCombo.SelectedIndex = DateTime.Today.Month - 1;
            PopulateDaysForCurrentMonth();
            UpdateSelectedDateLabel(DateTime.Today);
        }

        // Save edits currently present in the right-hand detail fields into the given DayViewModel
        private void SaveCurrentDayEdits()
        {
            try
            {
                if (_activeDay == null) return;
                // Persist text fields into DayViewModel overrides
                _activeDay.PositionOverride = PositionCombo.Text ?? string.Empty;
                _activeDay.LocationOverride = LocationCombo.Text ?? string.Empty;
                _activeDay.PhysicalLocationOverride = PhysicalLocationCombo.Text ?? string.Empty;

                // Persist to service so overrides survive application restarts
                TimeTrackingService.Instance.UpsertOverride(_activeDay.Date, _activeDay.PositionOverride, _activeDay.LocationOverride, _activeDay.PhysicalLocationOverride, _activeDay.TargetHours, _activeDay.DayType);

                // persist shifts for this date
                TimeTrackingService.Instance.UpsertShiftsForDate(_activeDay.Date, _activeDay.Shifts);

                // refresh combo lists in case new values were added
                UpdateComboLists();

                // DayType is already applied via DayTypeCombo_SelectionChanged when changed
                // Recompute cumulatives and update accounts so UI reflects changes
                RecomputeCumulatives();
                UpdateAccountsDisplay();
            }
            catch { }
        }

        // Recompute cumulative account balances for _currentDays in chronological order
        private void RecomputeCumulatives()
        {
            try
            {
                if (_currentDays == null || _currentDays.Count == 0) return;

                var acc = TimeTrackingService.Instance.GetAccountState();
                double runningTIL = acc.TILOffset;
                double runningHoliday = acc.HolidayOffset;

                // ensure days are sorted ascending
                var ordered = _currentDays.OrderBy(d => d.Date).ToList();
                foreach (var d in ordered)
                {
                    double dayWorked = 0;
                    if (d.Shifts != null && d.Shifts.Count > 0)
                        dayWorked = d.Shifts.Sum(s => s.Hours);
                    else if (d.Worked.HasValue)
                        dayWorked = d.Worked.Value;

                    double dayTarget = d.TargetHours ?? 0;
                    if (dayTarget <= 0 && d.Template != null)
                    {
                        int idx = ((int)d.Date.DayOfWeek + 6) % 7;
                        if (d.Template.HoursPerWeekday != null && d.Template.HoursPerWeekday.Length == 7)
                            dayTarget = d.Template.HoursPerWeekday[idx];
                    }

                    double dayDelta = dayWorked - (dayTarget > 0 ? dayTarget : 0);

                    if (d.DayType == DayType.WorkingDay)
                    {
                        runningTIL += dayDelta;
                    }
                    else if (d.DayType == DayType.Vacation)
                    {
                        runningHoliday -= 1.0;
                    }

                    d.CumulativeTIL = runningTIL;
                    d.CumulativeHoliday = runningHoliday;
                }
            }
            catch { }
        }

        private void UpdateSelectedDateLabel(DateTime d)
        {
            try
            {
                var tb = this.FindName("SelectedDateLabel") as TextBlock;
                if (tb != null)
                    tb.Text = d.ToString("dddd dd/MM/yyyy");
            }
            catch { }
        }

        private void MonthGridItem_Click(object sender, RoutedEventArgs e)
        {
            // handler attached to Border MouseLeftButtonUp in XAML - find DataContext
            if (e is System.Windows.Input.MouseButtonEventArgs mbe && mbe.Source is FrameworkElement fe)
            {
                if (fe.DataContext is DayViewModel vm)
                {
                    // persist edits for previous selection
                    SaveCurrentDayEdits();

                    // clear previous selection flags
                    foreach (var d in _currentDays) d.IsSelected = false;

                    DaysList.SelectedItem = vm;
                    vm.IsSelected = true;
                    _activeDay = vm;

                    // ensure layout updated and item scrolled into view
                    DaysList.UpdateLayout();
                    DaysList.ScrollIntoView(vm);
                    UpdateSelectedDateLabel(vm.Date);

                    MonthGrid.Items.Refresh();
                }
            }
        }

        private void DaysList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Before switching selection, save edits to previous active
            SaveCurrentDayEdits();

            if (DaysList.SelectedItem is DayViewModel vm)
            {
                // clear previous selection flags
                foreach (var d in _currentDays) d.IsSelected = false;
                vm.IsSelected = true;

                // populate right side details for selected day (basic)
                // prefer persisted overrides if present
                PositionCombo.Text = !string.IsNullOrWhiteSpace(vm.PositionOverride) ? vm.PositionOverride : (vm.Template?.JobDescription ?? "");
                LocationCombo.Text = !string.IsNullOrWhiteSpace(vm.LocationOverride) ? vm.LocationOverride : vm.Template?.Location ?? "";
                // default physical location to physical override, otherwise location override or template location
                PhysicalLocationCombo.Text = !string.IsNullOrWhiteSpace(vm.PhysicalLocationOverride) ? vm.PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(vm.LocationOverride) ? vm.LocationOverride : vm.Template?.Location ?? "");
                ShiftsList.ItemsSource = vm.Shifts;
                // Immediately set displayed total so it is visible even if UpdateShiftTotalsDisplay silently fails
                try
                {
                    double quickTotal = 0;
                    if (vm.Shifts != null && vm.Shifts.Count > 0)
                        quickTotal = vm.Shifts.Sum(s => s.Hours);
                    else if (vm.Worked.HasValue)
                        quickTotal = vm.Worked.Value;
                    TotalWorkedDataTextBlock.Text = quickTotal.ToString("0.##") + "h";
                }
                catch { }
                 DayTypeCombo.SelectedItem = vm.DayType.ToString();
                 UpdateSelectedDateLabel(vm.Date);

                 _activeDay = vm;

                 MonthGrid.Items.Refresh();

                 // Update shift totals and target display
                 UpdateShiftTotalsDisplay();
                 // Also refresh accounts display
                 UpdateAccountsDisplay();
            }
            else
            {
                _activeDay = null;
            }
        }

        private void ShiftsList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (_activeDay == null) return;
                if (ShiftsList.SelectedItem is Shift s)
                {
                    // reuse EditShift_Click logic by invoking same behavior
                    // clone for editing
                    var copy = new Shift { Date = s.Date, Start = s.Start, End = s.End, Description = s.Description, ManualStartOverride = s.ManualStartOverride, ManualEndOverride = s.ManualEndOverride, LunchBreak = s.LunchBreak, DayMode = s.DayMode };
                    var dlg = new ShiftDialog(copy, _activeDay.Date) { Owner = this };
                    if (dlg.ShowDialog() == true)
                    {
                        s.Start = dlg.Shift.Start;
                        s.End = dlg.Shift.End;
                        s.Description = dlg.Shift.Description;
                        s.ManualStartOverride = dlg.Shift.ManualStartOverride;
                        s.ManualEndOverride = dlg.Shift.ManualEndOverride;
                        s.LunchBreak = dlg.Shift.LunchBreak;
                        s.DayMode = dlg.Shift.DayMode;
                        // persist shifts
                        TimeTrackingService.Instance.UpsertShiftsForDate(_activeDay.Date, _activeDay.Shifts);
                        ShiftsList.Items.Refresh();
                        UpdateShiftTotalsDisplay();
                        RecomputeCumulatives();
                        UpdateAccountsDisplay();
                    }
                }
            }
            catch { }
        }

        private void AddShift_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_activeDay == null)
                {
                    MessageBox.Show("Please select a day first.", "Add Shift", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                var dlg = new ShiftDialog(null, _activeDay.Date) { Owner = this };
                if (dlg.ShowDialog() == true)
                {
                    _activeDay.Shifts.Add(dlg.Shift);
                    ShiftsList.Items.Refresh();
                    // persist shifts for this date
                    TimeTrackingService.Instance.UpsertShiftsForDate(_activeDay.Date, _activeDay.Shifts);
                    UpdateShiftTotalsDisplay();
                    RecomputeCumulatives();
                    UpdateAccountsDisplay();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to add shift: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void EditShift_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_activeDay == null)
                {
                    MessageBox.Show("Please select a day first.", "Edit Shift", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (ShiftsList.SelectedItem is Shift s)
                {
                    var copy = new Shift { Date = s.Date, Start = s.Start, End = s.End, Description = s.Description, ManualStartOverride = s.ManualStartOverride, ManualEndOverride = s.ManualEndOverride, LunchBreak = s.LunchBreak, DayMode = s.DayMode };
                    var dlg = new ShiftDialog(copy, _activeDay.Date) { Owner = this };
                    if (dlg.ShowDialog() == true)
                    {
                        s.Start = dlg.Shift.Start;
                        s.End = dlg.Shift.End;
                        s.Description = dlg.Shift.Description;
                        s.ManualStartOverride = dlg.Shift.ManualStartOverride;
                        s.ManualEndOverride = dlg.Shift.ManualEndOverride;
                        s.LunchBreak = dlg.Shift.LunchBreak;
                        s.DayMode = dlg.Shift.DayMode;
                        // persist shifts
                        TimeTrackingService.Instance.UpsertShiftsForDate(_activeDay.Date, _activeDay.Shifts);
                        ShiftsList.Items.Refresh();
                        UpdateShiftTotalsDisplay();
                        RecomputeCumulatives();
                        UpdateAccountsDisplay();
                    }
                }
                else
                {
                    MessageBox.Show("Please select a shift to edit.", "Edit Shift", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to edit shift: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DeleteShift_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_activeDay == null)
                {
                    MessageBox.Show("Please select a day first.", "Delete Shift", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (ShiftsList.SelectedItem is Shift s)
                {
                    if (MessageBox.Show($"Delete shift '{s.Display}'?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        _activeDay.Shifts.Remove(s);
                        ShiftsList.Items.Refresh();
                        TimeTrackingService.Instance.UpsertShiftsForDate(_activeDay.Date, _activeDay.Shifts);
                        UpdateShiftTotalsDisplay();
                        RecomputeCumulatives();
                        UpdateAccountsDisplay();
                    }
                }
                else
                {
                    MessageBox.Show("Please select a shift to delete.", "Delete Shift", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to delete shift: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ShiftTargetText_LostFocus(object sender, RoutedEventArgs e)
        {
            // on lost focus persist target
            if (_activeDay == null) return;
            if (this.FindName("ShiftTargetText") is TextBox tb && double.TryParse(tb.Text, out var val))
            {
                _activeDay.TargetHours = val;
                SaveCurrentDayEdits();
            }
            UpdateShiftTotalsDisplay();
        }

        private void SaveDefaults_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new TimeTemplateDialog() { Owner = this };
                if (dlg.ShowDialog() == true)
                {
                    // Save template
                    TimeTrackingService.Instance.UpsertTemplate(dlg.Template);
                    LoadTemplates();

                    // Directly apply the template across its date range, overwriting existing events and overrides
                    int applied = TimeTrackingService.Instance.ApplyTemplateWithOverrides(dlg.Template, overwriteExistingEvents: true, overwriteOverrides: true);
                    MessageBox.Show($"Bulk applied to {applied} day(s).", "Bulk apply", MessageBoxButton.OK, MessageBoxImage.Information);

                    // Refresh UI state
                    PopulateDaysForCurrentMonth();
                    RecomputeCumulatives();
                    UpdateAccountsDisplay();
                    UpdateComboLists();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save template: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ApplyDefaultRange_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Apply default to range - not yet implemented.");
        }

        private void ApplyDefaultFromDate_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Apply default from date - not yet implemented.");
        }

        private void ViewLog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new AccountLogDialog() { Owner = this };
                dlg.ShowDialog();
            }
            catch { }
        }

        private void ResetHoliday_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var acc = TimeTrackingService.Instance.GetAccountState();

                var dlg = new Window
                {
                    Title = "Reset Holiday account",
                    SizeToContent = SizeToContent.WidthAndHeight,
                    ResizeMode = ResizeMode.NoResize,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this
                };
                var stack = new StackPanel { Margin = new Thickness(12) };
                stack.Children.Add(new TextBlock { Text = "Enter new Holiday balance:" });
                var inputPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0,8,0,8) };
                var txt = new TextBox { Width = 160, Text = acc.HolidayOffset.ToString("0.##", CultureInfo.CurrentCulture) };
                inputPanel.Children.Add(txt);
                inputPanel.Children.Add(new TextBlock { Text = " days", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6,0,0,0) });
                stack.Children.Add(inputPanel);
                var chk = new CheckBox { Content = "Apply as relative delta (add to existing)", Margin = new Thickness(0,4,0,8) };
                stack.Children.Add(chk);
                var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
                var ok = new Button { Content = "OK", Width = 80, IsDefault = true, Margin = new Thickness(0,0,6,0) };
                var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
                btnPanel.Children.Add(ok); btnPanel.Children.Add(cancel); stack.Children.Add(btnPanel);
                dlg.Content = stack;

                ok.Click += (s, ea) =>
                {
                    if (double.TryParse(txt.Text, System.Globalization.NumberStyles.Any, CultureInfo.CurrentCulture, out var v))
                    {
                        dlg.Tag = new Tuple<double, bool>(v, chk.IsChecked == true);
                        dlg.DialogResult = true;
                    }
                    else
                        MessageBox.Show("Enter a valid number", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                };

                var res = dlg.ShowDialog();
                if (res == true && dlg.Tag is Tuple<double, bool> tup)
                {
                    var val = tup.Item1; var applyAsDelta = tup.Item2;
                    double old = acc.HolidayOffset; double delta;
                    if (applyAsDelta) { delta = val; acc.HolidayOffset = old + val; }
                    else { acc.HolidayOffset = val; delta = val - old; }
                    TimeTrackingService.Instance.SetAccountState(acc);
                    var entry = new AccountLogEntry { Date = DateTime.Now, Kind = "Holiday", Delta = delta, Balance = acc.HolidayOffset, Note = applyAsDelta ? "Manual delta" : "Manual set" };
                    TimeTrackingService.Instance.AddAccountLogEntry(entry);

                    // Reload persisted state to ensure any other components reading from disk see latest data
                    try { TimeTrackingService.Instance.Reload(); } catch { }

                    // Preserve current selection date
                    var selectedDate = _activeDay?.Date;

                    // Refresh computed cumulatives and UI so changes are visible immediately
                    try
                    {
                        PopulateDaysForCurrentMonth();
                        RecomputeCumulatives();
                        UpdateAccountsDisplay();

                        // also explicitly set account textblocks from service to ensure update
                        try
                        {
                            var fresh = TimeTrackingService.Instance.GetAccountState();
                            if (this.FindName("AccountHolidayValue") is TextBlock ah2)
                                ah2.Text = fresh.HolidayOffset.ToString("0.##") + " days";
                            if (this.FindName("AccountTILValue") is TextBlock at2)
                                at2.Text = fresh.TILOffset.ToString("0.##") + " hours";
                        }
                        catch { }

                        // restore selection if possible
                        if (selectedDate.HasValue)
                        {
                            var sel = _currentDays.FirstOrDefault(d => d.Date.Date == selectedDate.Value.Date);
                            if (sel != null)
                            {
                                DaysList.SelectedItem = sel;
                                sel.IsSelected = true;
                                _activeDay = sel;
                            }
                        }

                        DaysList.Items.Refresh();
                        MonthGrid.Items.Refresh();
                    }
                    catch { }

                    MessageBox.Show("Holiday account updated.", "Reset", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch { }
        }

        private void ResetTIL_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var acc = TimeTrackingService.Instance.GetAccountState();

                var dlg = new Window
                {
                    Title = "Reset Time in Lieu account",
                    SizeToContent = SizeToContent.WidthAndHeight,
                    ResizeMode = ResizeMode.NoResize,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this
                };
                var stack = new StackPanel { Margin = new Thickness(12) };
                stack.Children.Add(new TextBlock { Text = "Enter new Time in Lieu balance:" });
                var inputPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0,8,0,8) };
                var txt = new TextBox { Width = 160, Text = acc.TILOffset.ToString("0.##", CultureInfo.CurrentCulture) };
                inputPanel.Children.Add(txt);
                inputPanel.Children.Add(new TextBlock { Text = " hours", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6,0,0,0) });
                stack.Children.Add(inputPanel);
                var chk = new CheckBox { Content = "Apply as relative delta (add to existing)", Margin = new Thickness(0,4,0,8) };
                stack.Children.Add(chk);
                var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
                var ok = new Button { Content = "OK", Width = 80, IsDefault = true, Margin = new Thickness(0,0,6,0) };
                var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
                btnPanel.Children.Add(ok); btnPanel.Children.Add(cancel); stack.Children.Add(btnPanel);
                dlg.Content = stack;

                ok.Click += (s, ea) =>
                {
                    if (double.TryParse(txt.Text, System.Globalization.NumberStyles.Any, CultureInfo.CurrentCulture, out var v))
                    {
                        dlg.Tag = new Tuple<double, bool>(v, chk.IsChecked == true);
                        dlg.DialogResult = true;
                    }
                    else
                        MessageBox.Show("Enter a valid number", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                };

                var res = dlg.ShowDialog();
                if (res == true && dlg.Tag is Tuple<double, bool> tup)
                {
                    var val = tup.Item1; var applyAsDelta = tup.Item2;
                    double old = acc.TILOffset; double delta;
                    if (applyAsDelta) { delta = val; acc.TILOffset = old + val; }
                    else { acc.TILOffset = val; delta = val - old; }
                    TimeTrackingService.Instance.SetAccountState(acc);
                    var entry = new AccountLogEntry { Date = DateTime.Now, Kind = "TIL", Delta = delta, Balance = acc.TILOffset, Note = applyAsDelta ? "Manual delta" : "Manual set" };
                    TimeTrackingService.Instance.AddAccountLogEntry(entry);

                    // Reload persisted state
                    try { TimeTrackingService.Instance.Reload(); } catch { }

                    // preserve current selected date
                    var selectedDate = _activeDay?.Date;

                    // Refresh computed cumulatives and UI so changes are visible immediately
                    try
                    {
                        PopulateDaysForCurrentMonth();
                        RecomputeCumulatives();
                        UpdateAccountsDisplay();

                        // explicitly update account textblocks
                        try
                        {
                            var fresh = TimeTrackingService.Instance.GetAccountState();
                            if (this.FindName("AccountHolidayValue") is TextBlock ah3)
                                ah3.Text = fresh.HolidayOffset.ToString("0.##") + " days";
                            if (this.FindName("AccountTILValue") is TextBlock at3)
                                at3.Text = fresh.TILOffset.ToString("0.##") + " hours";
                        }
                        catch { }

                        if (selectedDate.HasValue)
                        {
                            var sel = _currentDays.FirstOrDefault(d => d.Date.Date == selectedDate.Value.Date);
                            if (sel != null)
                            {
                                DaysList.SelectedItem = sel;
                                sel.IsSelected = true;
                                _activeDay = sel;
                            }
                        }

                        DaysList.Items.Refresh();
                        MonthGrid.Items.Refresh();
                    }
                    catch { }

                    MessageBox.Show("Time in Lieu account updated.", "Reset", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch { }
        }

        // New helper: update totals display section when shifts/targets change
        private void UpdateShiftTotalsDisplay()
        {
            try
            {
                if (_activeDay == null)
                {
                    if (this.FindName("TotalWorkedDataTextBlock") is TextBlock tw) tw.Text = "0.00h";
                    if (this.FindName("TotalDeltaDataTextBlock") is TextBlock td) td.Text = "0.00h";
                    if (this.FindName("ShiftTargetText") is TextBox st) st.Text = string.Empty;
                    return;
                }

                double totalWorked = 0;
                if (_activeDay.Shifts != null && _activeDay.Shifts.Count > 0)
                    totalWorked = _activeDay.Shifts.Sum(s => s.Hours);
                else if (_activeDay.Worked.HasValue)
                    totalWorked = _activeDay.Worked.Value;

                double target = _activeDay.TargetHours ?? 0;
                if (target <= 0 && _activeDay.Template != null)
                {
                    int idx = ((int)_activeDay.Date.DayOfWeek + 6) % 7;
                    if (_activeDay.Template.HoursPerWeekday != null && _activeDay.Template.HoursPerWeekday.Length == 7)
                        target = _activeDay.Template.HoursPerWeekday[idx];
                }

                double delta = totalWorked - (target > 0 ? target : 0);

                if (this.FindName("TotalWorkedDataTextBlock") is TextBlock tw2)
                    tw2.Text = totalWorked.ToString("0.##") + "h";
                if (this.FindName("TotalDeltaDataTextBlock") is TextBlock td2)
                    td2.Text = delta.ToString("0.##") + "h";
                if (this.FindName("ShiftTargetText") is TextBox st2)
                    st2.Text = target > 0 ? target.ToString("0.##") : string.Empty;
            }
            catch { }
        }

        // New helper: update account displays (Holiday and TIL)
        private void UpdateAccountsDisplay()
        {
            try
            {
                double holidayVal, tilVal;
                // If an active day is selected, show cumulative balances as of the end of that day
                if (_activeDay != null)
                {
                    holidayVal = _activeDay.CumulativeHoliday;
                    tilVal = _activeDay.CumulativeTIL;
                }
                else
                {
                    var acc = TimeTrackingService.Instance.GetAccountState();
                    holidayVal = acc.HolidayOffset;
                    tilVal = acc.TILOffset;
                }

                if (this.FindName("AccountHolidayValue") is TextBlock ah)
                    ah.Text = holidayVal.ToString("0.##") + " days";
                if (this.FindName("AccountTILValue") is TextBlock til)
                    til.Text = tilVal.ToString("0.##") + " hours";

                 // Also refresh days list to show cumulative values
                 DaysList.Items.Refresh();
                 MonthGrid.Items.Refresh();
             }
             catch { }
         }

         // Handler for color toggle checkbox
         private void ColorModeCheck_Checked(object sender, RoutedEventArgs e)
         {
             ColorByPhysicalLocation = true;
             MonthGrid.Items.Refresh();
             DaysList.Items.Refresh();
         }

         private void ColorModeCheck_Unchecked(object sender, RoutedEventArgs e)
         {
             ColorByPhysicalLocation = false;
             MonthGrid.Items.Refresh();
             DaysList.Items.Refresh();
         }

        // Export year colour map button handler
        private void ExportYearMap_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (YearCombo.SelectedItem == null)
                {
                    MessageBox.Show("Select a year first.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                int year = int.Parse((string)YearCombo.SelectedItem);

                // gather all days for the year
                var first = new DateTime(year, 1, 1);
                var last = new DateTime(year, 12, 31);
                var summaries = TimeTrackingService.Instance.GetDaySummaries(first, last);
                var days = summaries.Select(s => new DayViewModel(s)).ToList();

                // apply overrides and shifts to days so DayType and PhysicalLocationOverride are accurate
                foreach (var d in days)
                {
                    var ov = TimeTrackingService.Instance.GetOverrideForDate(d.Date);
                    if (ov != null)
                    {
                        if (!string.IsNullOrWhiteSpace(ov.PhysicalLocation)) d.PhysicalLocationOverride = ov.PhysicalLocation;
                        else if (!string.IsNullOrWhiteSpace(ov.Location)) d.PhysicalLocationOverride = ov.Location;
                        d.PositionOverride = ov.Position; d.LocationOverride = ov.Location; d.TargetHours = ov.TargetHours; if (ov.DayType.HasValue) d.SetDayType(ov.DayType.Value);
                    }
                    // load persisted shifts
                    var saved = TimeTrackingService.Instance.GetShiftsForDate(d.Date).ToList();
                    if (saved.Any()) d.Shifts = saved.Select(s2 => new Shift { Date = s2.Date, Start = s2.Start, End = s2.End, Description = s2.Description, ManualStartOverride = s2.ManualStartOverride, ManualEndOverride = s2.ManualEndOverride, LunchBreak = s2.LunchBreak, DayMode = s2.DayMode }).ToList();
                }

                // prompt for filename
                var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "PDF files (*.pdf)|*.pdf", FileName = $"timemap_{year}.pdf" };
                if (dlg.ShowDialog() != true) return;

                // create PDF with two pages: Day type map and Physical location map
                using (var doc = new PdfDocument())
                {
                    // page 1: day type
                    var p1 = doc.AddPage();
                    p1.Size = PdfSharpCore.PageSize.A4;
                    using (var g = XGraphics.FromPdfPage(p1))
                    {
                        DrawYearGridToGraphics(g, days, year, colorByPhysical: false);
                    }

                    // page 2: physical location
                    var p2 = doc.AddPage();
                    p2.Size = PdfSharpCore.PageSize.A4;
                    using (var g = XGraphics.FromPdfPage(p2))
                    {
                        DrawYearGridToGraphics(g, days, year, colorByPhysical: true);
                    }

                    // page 3: overtime (delta) jet colour map
                    var p3 = doc.AddPage();
                    p3.Size = PdfSharpCore.PageSize.A4;
                    using (var g = XGraphics.FromPdfPage(p3))
                    {
                        DrawOvertimeGridToGraphics(g, days, year);
                    }

                    using (var fs = File.OpenWrite(dlg.FileName))
                    {
                        doc.Save(fs);
                    }
                }

                MessageBox.Show($"Exported year colour map to PDF.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to export: " + ex.Message, "Export", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Draws a monthly grid layout for the whole year on the given graphics surface.
        private void DrawYearGridToGraphics(XGraphics g, List<DayViewModel> days, int year, bool colorByPhysical)
        {
            // page layout constants
            double margin = 30;
            double pageWidth = g.PageSize.Width;
            double pageHeight = g.PageSize.Height;
            double contentWidth = pageWidth - (margin * 2);
            // reserve room at bottom for a horizontal legend
            double legendHeight = 70;

            // draw header
            var headerFont = new XFont("Arial", 18, XFontStyle.Bold);
            var subFont = new XFont("Arial", 10, XFontStyle.Regular);
            g.DrawString((colorByPhysical ? "Physical location map - " : "Day type map - ") + year.ToString(), headerFont, XBrushes.Black, new XRect(margin, 10, contentWidth, 30), XStringFormats.TopLeft);

            // layout 12 months in 3x4 grid
            int cols = 3; int rows = 4;
            double gap = 8;
            double monthWidth = (contentWidth - (gap * (cols - 1))) / cols;
            // leave extra space at bottom for horizontal legend
            double monthGridAvailableHeight = pageHeight - 80 - legendHeight - (gap * (rows - 1));
            double monthHeight = monthGridAvailableHeight / rows;

            for (int m = 1; m <= 12; m++)
            {
                int monthIndex = m - 1;
                int col = monthIndex % cols;
                int row = monthIndex / cols;
                double x = margin + col * (monthWidth + gap);
                double y = 50 + row * (monthHeight + gap);

                DrawSingleMonth(g, days, year, m, new XRect(x, y, monthWidth, monthHeight), colorByPhysical);
            }

            // add small legend horizontally across bottom
            DrawLegend(g, margin, pageHeight - legendHeight + 12, colorByPhysical, days, contentWidth);
        }

        private void DrawSingleMonth(XGraphics g, List<DayViewModel> days, int year, int month, XRect rect, bool colorByPhysical)
        {
            // draw month title
            var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(month);
            var titleFont = new XFont("Arial", 12, XFontStyle.Bold);
            g.DrawString(monthName + " " + year.ToString(), titleFont, XBrushes.Black, new XPoint(rect.X + 4, rect.Y + 14));

            // cell grid for 7 columns x up to 6 rows
            double top = rect.Y + 22;
            double left = rect.X + 2;
            double gridWidth = rect.Width - 4;
            double gridHeight = rect.Height - 28;
            int cols = 7;
            int rows = 6;
            double cellW = gridWidth / cols;
            double cellH = gridHeight / rows;

            // first day offset: start week on Monday
            var first = new DateTime(year, month, 1);
            int leading = (int)first.DayOfWeek - 1; if (leading < 0) leading += 7;

            for (int i = 0; i < rows * cols; i++)
            {
                int dayNum = i - leading + 1;
                double cx = left + (i % cols) * cellW;
                double cy = top + (i / cols) * cellH;
                var cellRect = new XRect(cx, cy, cellW - 1, cellH - 1);

                if (dayNum >= 1 && dayNum <= DateTime.DaysInMonth(year, month))
                {
                    var d = days.FirstOrDefault(dd => dd.Date.Year == year && dd.Date.Month == month && dd.Date.Day == dayNum);
                    XBrush brush = XBrushes.White;
                    if (d != null)
                    {
                        if (colorByPhysical)
                        {
                            var key = !string.IsNullOrWhiteSpace(d.PhysicalLocationOverride) ? d.PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(d.LocationOverride) ? d.LocationOverride : d.Template?.Location ?? "");
                            if (string.IsNullOrWhiteSpace(key)) brush = XBrushes.White;
                            else
                            {
                                int h = key.GetHashCode();
                                byte r = (byte)(80 + (Math.Abs(h) % 176));
                                byte gcol = (byte)(80 + (Math.Abs(h / 7) % 176));
                                byte b = (byte)(80 + (Math.Abs(h / 13) % 176));
                                brush = new XSolidBrush(XColor.FromArgb(255, r, gcol, b));
                            }
                        }
                        else
                        {
                            if (DayTypePdfColors.TryGetValue(d.DayType, out var c)) brush = new XSolidBrush(c);
                            else brush = XBrushes.White;
                        }
                    }

                    g.DrawRectangle(brush, cellRect);
                    // draw day number
                    var numFont = new XFont("Arial", 8, XFontStyle.Regular);
                    g.DrawString(dayNum.ToString(), numFont, XBrushes.Black, new XPoint(cellRect.X + 4, cellRect.Y + 10));
                }
                else
                {
                    // blank
                    g.DrawRectangle(XBrushes.LightGray, cellRect);
                }
            }

            // border
            g.DrawRectangle(XPens.Black, rect.X, rect.Y, rect.Width, rect.Height);
        }

        private void DrawLegend(XGraphics g, double x, double y, bool colorByPhysical, List<DayViewModel> days, double availableWidth)
        {
            var font = new XFont("Arial", 9, XFontStyle.Regular);
            double box = 12; double gap = 8; double cx = x;
            double startX = x;
            double padding = 6;
            // compute available horizontal area and center legend items
            if (!colorByPhysical)
            {
                var items = DayTypePdfColors.Keys.Select(k => k.ToString()).ToList();
                // estimate width per item
                var widths = items.Select(it => g.MeasureString(it, font).Width + box + gap + padding).ToList();
                double totalW = widths.Sum();
                cx = startX + Math.Max(0, (availableWidth - totalW) / 2);
                for (int i = 0; i < items.Count; i++)
                {
                    var key = items[i];
                    var brush = new XSolidBrush(DayTypePdfColors[(DayType)Enum.Parse(typeof(DayType), key)]);
                    g.DrawRectangle(brush, cx, y, box, box);
                    g.DrawString(key, font, XBrushes.Black, new XPoint(cx + box + 6, y + box - 2));
                    cx += widths[i];
                }
            }
            else
            {
                // show top N physical locations found horizontally
                var locs = daysForLegendLookup();
                int max = Math.Min(8, locs.Count);
                var widths = new List<double>();
                for (int i = 0; i < max; i++) widths.Add(g.MeasureString(locs[i], font).Width + box + gap + padding);
                double totalW = widths.Sum();
                cx = startX + Math.Max(0, (availableWidth - totalW) / 2);
                for (int i = 0; i < max; i++)
                {
                    var key = locs[i];
                    int h = key.GetHashCode();
                    byte r = (byte)(80 + (Math.Abs(h) % 176));
                    byte gcol = (byte)(80 + (Math.Abs(h / 7) % 176));
                    byte b = (byte)(80 + (Math.Abs(h / 13) % 176));
                    var brush = new XSolidBrush(XColor.FromArgb(255, r, gcol, b));
                    g.DrawRectangle(brush, cx, y, box, box);
                    g.DrawString(key, font, XBrushes.Black, new XPoint(cx + box + 6, y + box - 2));
                    cx += widths[i];
                }
            }

            // local function to obtain most common physical locations
            List<string> daysForLegendLookup()
            {
                var locs = days.Where(d => !string.IsNullOrWhiteSpace(d.PhysicalLocationOverride) || !string.IsNullOrWhiteSpace(d.LocationOverride) || !string.IsNullOrWhiteSpace(d.Template?.Location))
                               .Select(d => !string.IsNullOrWhiteSpace(d.PhysicalLocationOverride) ? d.PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(d.LocationOverride) ? d.LocationOverride : d.Template?.Location ?? ""))
                               .Where(s => !string.IsNullOrWhiteSpace(s))
                               .GroupBy(s => s).OrderByDescending(gp => gp.Count()).Select(gp => gp.Key).ToList();
                return locs;
            }
        }

        // Third page: draw overtime (delta) map using a 'jet' style colormap
        private void DrawOvertimeGridToGraphics(XGraphics g, List<DayViewModel> days, int year)
        {
            double margin = 30;
            double pageWidth = g.PageSize.Width;
            double pageHeight = g.PageSize.Height;
            double contentWidth = pageWidth - (margin * 2);
            
            var headerFont = new XFont("Arial", 16, XFontStyle.Bold);
            var subFont = new XFont("Arial", 10, XFontStyle.Regular);

            // compute deltas
            // Fixed scale +/- 8 hours as requested
            double min = -8.0;
            double max = 8.0;

            double totalOver = days.Where(d => d.Delta.HasValue && d.Delta.Value > 0).Sum(d => d.Delta.Value);
            int countOver = days.Count(d => d.Delta.HasValue && d.Delta.Value > 0);

            g.DrawString("Overtime (Delta) map - " + year.ToString(), headerFont, XBrushes.Black, new XRect(margin, 10, contentWidth, 24), XStringFormats.TopLeft);
            g.DrawString($"Total overtime: {totalOver:0.##}h across {countOver} day(s)", subFont, XBrushes.Black, new XRect(margin, 34, contentWidth, 20), XStringFormats.TopLeft);

            // layout months grid (3x4)
            int cols = 3; int rows = 4; double gap = 8;
            double monthWidth = (contentWidth - (gap * (cols - 1))) / cols;
            double legendHeight = 70;
            double monthGridAvailableHeight = pageHeight - 80 - legendHeight - (gap * (rows - 1));
            double monthHeight = monthGridAvailableHeight / rows;

            for (int m = 1; m <= 12; m++)
            {
                int monthIndex = m - 1;
                int col = monthIndex % cols;
                int row = monthIndex / cols;
                double x = margin + col * (monthWidth + gap);
                double y = 60 + row * (monthHeight + gap);

                // Draw similar month grid but color by delta
                DrawSingleMonthOvertime(g, days, year, m, new XRect(x, y, monthWidth, monthHeight), min, max);
            }

            // draw horizontal colorbar legend at bottom
            double legendY = pageHeight - 60;
            double legendX = margin + 20;
            double legendW = contentWidth - 40;
            double legendH = 14;
            int steps = Math.Max(4, (int)legendW); // ensure at least a few steps
            for (int i = 0; i < steps; i++)
            {
                double t = (double)i / Math.Max(1, steps - 1);
                double val = min + t * (max - min);
                var c = JetColor(val, min, max);
                var brush = new XSolidBrush(c);
                g.DrawRectangle(brush, legendX + i, legendY, 1, legendH);
            }
            // draw labels min 0 max
            var lblFont = new XFont("Arial", 9, XFontStyle.Regular);
            g.DrawString(min.ToString("0.##") + "h", lblFont, XBrushes.Black, new XPoint(legendX, legendY + legendH + 6));
            // safe position for zero marker
            double zeroPos = legendX;
            if (max > min) zeroPos = legendX + legendW * ((0 - min) / (max - min));
            g.DrawString("0h", lblFont, XBrushes.Black, new XPoint(zeroPos - g.MeasureString("0h", lblFont).Width/2, legendY + legendH + 6));
            g.DrawString(max.ToString("0.##") + "h", lblFont, XBrushes.Black, new XPoint(legendX + legendW - g.MeasureString(max.ToString("0.##") + "h", lblFont).Width, legendY + legendH + 6));
        }

        private void DrawSingleMonthOvertime(XGraphics g, List<DayViewModel> days, int year, int month, XRect rect, double min, double max)
        {
            var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(month);
            var titleFont = new XFont("Arial", 10, XFontStyle.Bold);
            g.DrawString(monthName + " " + year.ToString(), titleFont, XBrushes.Black, new XPoint(rect.X + 4, rect.Y + 12));

            double top = rect.Y + 20;
            double left = rect.X + 2;
            double gridWidth = rect.Width - 4;
            double gridHeight = rect.Height - 24;
            int cols = 7; int rows = 6;
            double cellW = gridWidth / cols;
            double cellH = gridHeight / rows;

            var first = new DateTime(year, month, 1);
            int leading = (int)first.DayOfWeek - 1; if (leading < 0) leading += 7;

            for (int i = 0; i < rows * cols; i++)
            {
                int dayNum = i - leading + 1;
                double cx = left + (i % cols) * cellW;
                double cy = top + (i / cols) * cellH;
                var cellRect = new XRect(cx, cy, cellW - 1, cellH - 1);

                if (dayNum >= 1 && dayNum <= DateTime.DaysInMonth(year, month))
                {
                    var d = days.FirstOrDefault(dd => dd.Date.Year == year && dd.Date.Month == month && dd.Date.Day == dayNum);
                    if (d != null && d.Delta.HasValue)
                    {
                        var jc = JetColor(d.Delta.Value, min, max);
                        g.DrawRectangle(new XSolidBrush(jc), cellRect);
                    }
                    else if (d != null)
                    {
                        g.DrawRectangle(XBrushes.White, cellRect);
                    }
                    else
                    {
                        g.DrawRectangle(XBrushes.LightGray, cellRect);
                    }

                    var numFont = new XFont("Arial", 7, XFontStyle.Regular);
                    g.DrawString(dayNum.ToString(), numFont, XBrushes.Black, new XPoint(cellRect.X + 3, cellRect.Y + 9));
                }
                else
                {
                    g.DrawRectangle(XBrushes.LightGray, cellRect);
                }
            }

            g.DrawRectangle(XPens.Black, rect.X, rect.Y, rect.Width, rect.Height);
        }

        // Map a value to a jet-like color between min and max. Returns XColor.
        private XColor JetColor(double value, double min, double max)
        {
            // Values outside [min,max] should be rendered in black
            if (value < min || value > max)
                return XColor.FromArgb(255, 0, 0, 0);

            double t = 0.5; // normalized
            if (max > min)
                t = (value - min) / (max - min);
            // clamp
            t = Math.Min(1, Math.Max(0, t));

            double r = 0, g = 0, b = 0;
            if (t < 0.25)
            {
                r = 0; g = 4 * t; b = 1;
            }
            else if (t < 0.5)
            {
                r = 0; g = 1; b = 1 - 4 * (t - 0.25);
            }
            else if (t < 0.75)
            {
                r = 4 * (t - 0.5); g = 1; b = 0;
            }
            else
            {
                r = 1; g = 1 - 4 * (t - 0.75); b = 0;
            }

            byte R = (byte)(Math.Min(1, Math.Max(0, r)) * 255);
            byte G = (byte)(Math.Min(1, Math.Max(0, g)) * 255);
            byte B = (byte)(Math.Min(1, Math.Max(0, b)) * 255);
            return XColor.FromArgb(255, R, G, B);
        }
    }

    // Simple settings for window position/size
    internal class WindowSettings
    {
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public WindowState State { get; set; }
    }

    public class DayViewModel
    {
        public DateTime Date { get; }
        public string DateString => Date.ToString("dd/MM/yyyy");
        public string Weekday => Date.ToString("dddd");
        public double? Worked { get; }
        public double Standard { get; }
        // cumulative account balances as of end of this day
        public double CumulativeTIL { get; set; }
        public double CumulativeHoliday { get; set; }
        public string CumulativeTILDisplay => CumulativeTIL.ToString("0.##") + "h";
        public string CumulativeHolidayDisplay => CumulativeHoliday.ToString("0.##") + "d";
        public double? Delta => Worked.HasValue ? (Worked.Value - Standard) : (double?)null;
        public string WorkedDisplay => Worked.HasValue ? Worked.Value.ToString("0.00") : "0:00";
        public string StandardDisplay => Standard.ToString("0.00");
        public DayType DayType { get; private set; }
        public TimeTemplate Template { get; }

        // New properties to indicate today and selection for XAML triggers
        public bool IsToday { get; set; }
        public bool IsSelected { get; set; }

        // Per-day overrides (user edits) to preserve values while dialog is open
        public string PositionOverride { get; set; }
        public string LocationOverride { get; set; }
        public string PhysicalLocationOverride { get; set; }

        // Per-day shifts (persisted)
        public List<Shift> Shifts { get; set; }

        // Target hours for the day (default from template, editable in dialog)
        public double? TargetHours { get; set; }

        public DayViewModel(DaySummary s)
        {
            Date = s.Date;
            Worked = s.WorkedHours;
            Standard = s.StandardHours;
            Template = TimeTrackingService.Instance.GetTemplates().FirstOrDefault(t => t.AppliesTo(s.Date));
            // determine type: weekend = Saturday or Sunday
            if (Date.DayOfWeek == DayOfWeek.Saturday || Date.DayOfWeek == DayOfWeek.Sunday)
                DayType = DayType.Weekend;
            else
                DayType = DayType.WorkingDay;

            IsToday = false;
            IsSelected = false;

            PositionOverride = null;
            LocationOverride = null;
            PhysicalLocationOverride = null;

            // load persisted shifts
            var saved = TimeTrackingService.Instance.GetShiftsForDate(Date).ToList();
            if (saved != null && saved.Count > 0)
            {
                Shifts = saved.Select(s2 => new Shift { Date = s2.Date, Start = s2.Start, End = s2.End, Description = s2.Description, ManualStartOverride = s2.ManualStartOverride, ManualEndOverride = s2.ManualEndOverride, LunchBreak = s2.LunchBreak, DayMode = s2.DayMode }).ToList();
            }
            else
            {
                Shifts = new List<Shift>();
            }

            // If no persisted shifts but DaySummary contains open/close times, create a synthetic shift so the UI shows what was used to compute WorkedHours
            if ((Shifts == null || Shifts.Count == 0) && s.OpenTime.HasValue && s.CloseTime.HasValue)
            {
                try
                {
                    var start = s.OpenTime.Value;
                    var end = s.CloseTime.Value;
                    // ensure valid ordering
                    if (end <= start)
                    {
                        // if close <= open, treat worked as 0 and do not add
                    }
                    else
                    {
                        var synthetic = new Shift
                        {
                            Date = Date,
                            Start = start,
                            End = end,
                            Description = "(from events)",
                            LunchBreak = TimeSpan.Zero,
                            ManualStartOverride = false,
                            ManualEndOverride = false,
                            DayMode = "auto"
                        };
                        Shifts.Add(synthetic);
                    }
                }
                catch { }
            }
        }

        public void SetDayType(DayType dt) => DayType = dt;

        // Computed background brush used by month grid. Respects global toggle in TimeTrackingDialog.ColorByPhysicalLocation
        public Brush BackgroundBrush
        {
            get
            {
                try
                {
                    if (TimeTrackingDialog.ColorByPhysicalLocation)
                    {
                        var key = !string.IsNullOrWhiteSpace(PhysicalLocationOverride) ? PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(LocationOverride) ? LocationOverride : Template?.Location ?? "");
                        if (string.IsNullOrWhiteSpace(key))
                            return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFFFFF"));

                        int h = key.GetHashCode();
                        // produce a pleasant pastel color
                        byte r = (byte)(80 + (Math.Abs(h) % 176));
                        byte g = (byte)(80 + (Math.Abs(h / 7) % 176));
                        byte b = (byte)(80 + (Math.Abs(h / 13) % 176));
                        return new SolidColorBrush(Color.FromRgb(r, g, b));
                    }
                    else
                    {
                        // colour by day type using existing palette
                        switch (DayType)
                        {
                            case DayType.WorkingDay: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFDFF0D8"));
                            case DayType.Weekend: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF0F0F0"));
                            case DayType.PublicHoliday: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFE5E5"));
                            case DayType.Vacation: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFDDEEFF"));
                            case DayType.TimeInLieu: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFF2CC"));
                            case DayType.Other: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFEFEFEF"));
                            default: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFFFFF"));
                        }
                    }
                }
                catch
                {
                    return new SolidColorBrush(Colors.White);
                }
            }
        }
    }
 }

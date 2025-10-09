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
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using Microsoft.Win32;
using ClosedXML.Excel;

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
                    var j = JsonSerializer.Deserialize<WindowSettings>(File.ReadAllText(WindowSettingsFile));
                    if (j != null)
                    {
                        this.WindowStartupLocation = WindowStartupLocation.Manual;
                        this.Left = j.Left;
                        this.Top = j.Top;
                        this.Width = j.Width;
                        this.Height = j.Height;
                        if (j.State == WindowState.Maximized)
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

            // Include carry-forward from all prior days before the first of this month so TIL does not reset monthly
            try
            {
                runningTIL += ComputeTILCarryForward(first);
            }
            catch { }

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
                else
                {
                    d.Shifts = new List<Shift>();
                }

                // compute worked hours for this day strictly from shifts
                double dayWorked = 0;
                if (d.Shifts != null && d.Shifts.Count > 0)
                    dayWorked = d.Shifts.Sum(s => s.Hours);

                // determine target hours for this day (override or template default)
                double dayTarget;
                if (d.TargetHours.HasValue)
                {
                    dayTarget = d.TargetHours.Value; // explicit override (can be zero)
                }
                else if (d.DayType == DayType.Vacation || d.DayType == DayType.PublicHoliday)
                {
                    dayTarget = 0; // auto-zero on vacation or public holiday when no explicit override
                }
                else if (d.Template != null)
                {
                    int idx = ((int)d.Date.DayOfWeek + 6) % 7;
                    if (d.Template.HoursPerWeekday != null && d.Template.HoursPerWeekday.Length == 7)
                        dayTarget = d.Template.HoursPerWeekday[idx];
                    else
                        dayTarget = 0;
                }
                else
                {
                    dayTarget = 0;
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
            int leading = ((int)first.DayOfWeek - 1); if (leading < 0) leading += 7;
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
                UpdateOverridePanels();
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

                        // If day marked as holiday or public holiday, set target hours to zero and persist override
                        if (dt == DayType.Vacation || dt == DayType.PublicHoliday)
                        {
                            vm.TargetHours = 0;
                            try { TimeTrackingService.Instance.UpsertOverride(vm.Date, vm.PositionOverride, vm.LocationOverride, vm.PhysicalLocationOverride, vm.TargetHours, dt); } catch { }
                            // update target textbox immediately
                            try { if (this.FindName("ShiftTargetText") is TextBox tb) tb.Text = "0"; } catch { }
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
                                    // Apply the log entry on the selected day rather than the current date/time
                                    Date = vm.Date.AddHours(12),
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
                        UpdateShiftTotalsDisplay();
                        UpdateOverridePanels();
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

        // Prev/Next month navigation buttons
        private void PrevMonthButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (YearCombo.SelectedItem is string ys && int.TryParse(ys, out var year))
                {
                    int monthIndex = MonthCombo.SelectedIndex; // 0..11
                    if (monthIndex > 0)
                    {
                        MonthCombo.SelectedIndex = monthIndex - 1;
                    }
                    else
                    {
                        int targetYear = year - 1;
                        EnsureYearInCombo(targetYear);
                        YearCombo.SelectedItem = targetYear.ToString();
                        MonthCombo.SelectedIndex = 11; // December
                    }
                }
            }
            catch { }
        }

        private void NextMonthButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (YearCombo.SelectedItem is string ys && int.TryParse(ys, out var year))
                {
                    int monthIndex = MonthCombo.SelectedIndex; // 0..11
                    if (monthIndex < 11)
                    {
                        MonthCombo.SelectedIndex = monthIndex + 1;
                    }
                    else
                    {
                        int targetYear = year + 1;
                        EnsureYearInCombo(targetYear);
                        YearCombo.SelectedItem = targetYear.ToString();
                        MonthCombo.SelectedIndex = 0; // January
                    }
                }
            }
            catch { }
        }

        private void EnsureYearInCombo(int year)
        {
            try
            {
                string ys = year.ToString();
                bool exists = false;
                foreach (var item in YearCombo.Items)
                {
                    if (item is string s && s == ys) { exists = true; break; }
                }
                if (!exists)
                {
                    // Insert keeping ascending order
                    int insertAt = 0;
                    for (int i = 0; i < YearCombo.Items.Count; i++)
                    {
                        if (int.TryParse(YearCombo.Items[i] as string, out var val) && val < year)
                            insertAt = i + 1;
                    }
                    YearCombo.Items.Insert(insertAt, ys);
                }
            }
            catch { }
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
                UpdateOverridePanels();
            }
            catch { }
        }

        // Compute TIL carry-forward from the earliest saved data up to (but not including) the given start date
        private double ComputeTILCarryForward(DateTime startDateExclusive)
        {
            double carry = 0.0;
            try
            {
                // determine earliest date we have data for (shifts or overrides)
                DateTime? minShift = null;
                try
                {
                    var allShifts = TimeTrackingService.Instance.GetAllSavedShifts();
                    if (allShifts != null && allShifts.Count > 0)
                        minShift = allShifts.Min(s => s.Date.Date);
                }
                catch { }

                DateTime? minOverride = null;
                try
                {
                    var allOverrides = TimeTrackingService.Instance.GetOverrides();
                    if (allOverrides != null && allOverrides.Count > 0)
                        minOverride = allOverrides.Min(o => o.Date.Date);
                }
                catch { }

                DateTime? start = null;
                if (minShift.HasValue && minOverride.HasValue) start = (minShift.Value < minOverride.Value) ? minShift : minOverride;
                else if (minShift.HasValue) start = minShift;
                else if (minOverride.HasValue) start = minOverride;

                if (!start.HasValue) return 0.0; // nothing to carry

                var end = startDateExclusive.Date.AddDays(-1);
                if (end < start.Value.Date) return 0.0;

                var templates = TimeTrackingService.Instance.GetTemplates();

                for (var d = start.Value.Date; d <= end; d = d.AddDays(1))
                {
                    // Determine day type (default weekend/working day unless overridden)
                    var dayType = (d.DayOfWeek == DayOfWeek.Saturday || d.DayOfWeek == DayOfWeek.Sunday) ? DayType.Weekend : DayType.WorkingDay;
                    var ov = TimeTrackingService.Instance.GetOverrideForDate(d);
                    if (ov != null && ov.DayType.HasValue) dayType = ov.DayType.Value;

                    // Sum worked hours from saved shifts
                    double worked = 0.0;
                    try
                    {
                        var sh = TimeTrackingService.Instance.GetShiftsForDate(d).ToList();
                        if (sh != null && sh.Count > 0)
                            worked = sh.Sum(s => s.Hours);
                    }
                    catch { }

                    // Determine target hours
                    double target;
                    if (ov != null && ov.TargetHours.HasValue)
                    {
                        target = ov.TargetHours.Value;
                    }
                    else if (dayType == DayType.Vacation || dayType == DayType.PublicHoliday)
                    {
                        target = 0;
                    }
                    else
                    {
                        double t = 0;
                        var temp = templates.FirstOrDefault(tpl => tpl.AppliesTo(d));
                        if (temp != null && temp.HoursPerWeekday != null && temp.HoursPerWeekday.Length == 7)
                        {
                            int idx = ((int)d.DayOfWeek + 6) % 7; // Monday=0
                            t = temp.HoursPerWeekday[idx];
                        }
                        target = t;
                    }

                    double delta = worked - (target > 0 ? target : 0);
                    if (dayType == DayType.WorkingDay)
                        carry += delta;
                    // Holiday days affect holiday account, not TIL; other types ignore for TIL accumulation
                }
            }
            catch { }
            return carry;
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

                // Include carry-forward from before the first day currently shown
                try
                {
                    var monthStart = _currentDays.Min(d => d.Date.Date);
                    runningTIL += ComputeTILCarryForward(monthStart);
                }
                catch { }

                // ensure days are sorted ascending
                var ordered = _currentDays.OrderBy(d => d.Date).ToList();
                foreach (var d in ordered)
                {
                    double dayWorked = 0;
                    if (d.Shifts != null && d.Shifts.Count > 0)
                        dayWorked = d.Shifts.Sum(s => s.Hours);
                    // no fallback to summary; 0 if no shifts

                    double dayTarget;
                    if (d.TargetHours.HasValue)
                    {
                        dayTarget = d.TargetHours.Value;
                    }
                    else if (d.DayType == DayType.Vacation || d.DayType == DayType.PublicHoliday)
                    {
                        dayTarget = 0;
                    }
                    else if (d.Template != null)
                    {
                        int idx = ((int)d.Date.DayOfWeek + 6) % 7;
                        if (d.Template.HoursPerWeekday != null && d.Template.HoursPerWeekday.Length == 7)
                            dayTarget = d.Template.HoursPerWeekday[idx];
                        else
                            dayTarget = 0;
                    }
                    else
                    {
                        dayTarget = 0;
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
                PhysicalLocationCombo.Text = !string.IsNullOrWhiteSpace(vm.PhysicalLocationOverride) ? vm.PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(vm.LocationOverride) ? vm.LocationOverride : vm.Template?.Location ?? "");
                ShiftsList.ItemsSource = vm.Shifts;
                // Immediately set displayed total so it is visible even if UpdateShiftTotalsDisplay silently fails
                try
                {
                    double quickTotal = 0;
                    if (vm.Shifts != null && vm.Shifts.Count > 0)
                        quickTotal = vm.Shifts.Sum(s => s.Hours);
                    // no fallback to summary
                    TotalWorkedDataTextBlock.Text = FormatHoursAsHHmm(quickTotal);
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
                 // Update manual override indicator panels
                 UpdateOverridePanels();
            }
            else
            {
                _activeDay = null;
                UpdateOverridePanels();
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
                        UpdateOverridePanels();
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
                    UpdateOverridePanels();
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
                        UpdateOverridePanels();
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
                        UpdateOverridePanels();
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
            if (this.FindName("ShiftTargetText") is TextBox tb)
            {
                var text = tb.Text?.Trim();
                if (string.IsNullOrEmpty(text))
                {
                    // empty clears override
                    _activeDay.TargetHours = null;
                }
                else if (double.TryParse(text, out var val))
                {
                    _activeDay.TargetHours = val; // allow explicit zero
                }
            }
            SaveCurrentDayEdits();
            UpdateShiftTotalsDisplay();
            UpdateOverridePanels();
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
                    UpdateShiftTotalsDisplay();
                    UpdateOverridePanels();
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
                // refresh after closing log in case user changed anything there in future
                UpdateOverridePanels();
            }
            catch { }
        }

        private void EditLocationColors_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new LocationColorsDialog() { Owner = this };
                if (dlg.ShowDialog() == true)
                {
                    // refresh UI to apply new colors
                    MonthGrid.Items.Refresh();
                    DaysList.Items.Refresh();
                }
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

                    // Apply the manual change on the selected day in the viewer
                    var targetDate = _activeDay?.Date.Date ?? DateTime.Today.Date;
                    var entry = new AccountLogEntry { Date = targetDate.AddHours(12), Kind = "Holiday", Delta = delta, Balance = acc.HolidayOffset, Note = applyAsDelta ? "Manual delta" : "Manual set", AffectedDate = targetDate };
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
                        UpdateShiftTotalsDisplay();
                        UpdateOverridePanels();

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

                    // Apply the manual change on the selected day in the viewer
                    var targetDate = _activeDay?.Date.Date ?? DateTime.Today.Date;
                    var entry = new AccountLogEntry { Date = targetDate.AddHours(12), Kind = "TIL", Delta = delta, Balance = acc.TILOffset, Note = applyAsDelta ? "Manual delta" : "Manual set", AffectedDate = targetDate };
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
                        UpdateShiftTotalsDisplay();
                        UpdateOverridePanels();

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

        // New: manual override summary and edit/delete controls
        private void UpdateOverridePanels()
        {
            try
            {
                var holPanel = this.FindName("HolidayOverridePanel") as FrameworkElement;
                var holLabel = this.FindName("HolidayOverrideLabel") as TextBlock;
                var tilPanel = this.FindName("TILOverridePanel") as FrameworkElement;
                var tilLabel = this.FindName("TILOverrideLabel") as TextBlock;

                if (_activeDay == null)
                {
                    if (holPanel != null) holPanel.Visibility = Visibility.Collapsed;
                    if (tilPanel != null) tilPanel.Visibility = Visibility.Collapsed;
                    return;
                }

                var date = _activeDay.Date.Date;
                double hol = GetManualOverrideSumForDay(date, "Holiday");
                double til = GetManualOverrideSumForDay(date, "TIL");

                if (holPanel != null && holLabel != null)
                {
                    if (Math.Abs(hol) > 0.0001)
                    {
                        holPanel.Visibility = Visibility.Visible;
                        holLabel.Text = $"Manual override: {hol:+0.##;-0.##;0} d";
                    }
                    else holPanel.Visibility = Visibility.Collapsed;
                }
                if (tilPanel != null && tilLabel != null)
                {
                    if (Math.Abs(til) > 0.0001)
                    {
                        tilPanel.Visibility = Visibility.Visible;
                        tilLabel.Text = $"Manual override: {FormatHoursAsHHmm(til)}";
                    }
                    else tilPanel.Visibility = Visibility.Collapsed;
                }
            }
            catch { }
        }

        private static double GetManualOverrideSumForDay(DateTime date, string kind)
        {
            try
            {
                var items = TimeTrackingService.Instance.GetAccountLogEntries(date, date, kind);
                if (items == null) return 0;
                return items.Where(a => (a.Note ?? string.Empty).IndexOf("manual", StringComparison.OrdinalIgnoreCase) >= 0)
                            .Sum(a => a.Delta);
            }
            catch { return 0; }
        }

        private void EditTILOverride_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_activeDay == null) return;
                var current = GetManualOverrideSumForDay(_activeDay.Date.Date, "TIL");

                var dlg = new Window
                {
                    Title = "Edit Time in Lieu manual override",
                    SizeToContent = SizeToContent.WidthAndHeight,
                    ResizeMode = ResizeMode.NoResize,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this
                };
                var stack = new StackPanel { Margin = new Thickness(12) };
                stack.Children.Add(new TextBlock { Text = "Enter total manual override for this day:" });
                var inputPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0,8,0,8) };
                var txt = new TextBox { Width = 160, Text = current.ToString("0.##", CultureInfo.CurrentCulture) };
                inputPanel.Children.Add(txt);
                inputPanel.Children.Add(new TextBlock { Text = " hours", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6,0,0,0) });
                stack.Children.Add(inputPanel);
                var btns = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
                var ok = new Button { Content = "OK", Width = 80, IsDefault = true, Margin = new Thickness(0,0,6,0) };
                var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
                btns.Children.Add(ok); btns.Children.Add(cancel); stack.Children.Add(btns);
                dlg.Content = stack;

                ok.Click += (s, ea) =>
                {
                    if (double.TryParse(txt.Text, NumberStyles.Any, CultureInfo.CurrentCulture, out var v)) { dlg.Tag = v; dlg.DialogResult = true; }
                    else MessageBox.Show("Enter a valid number", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                };

                if (dlg.ShowDialog() == true)
                {
                    var desired = (double)dlg.Tag;
                    var diff = desired - current;
                    if (Math.Abs(diff) > 0.0001)
                    {
                        var acc = TimeTrackingService.Instance.GetAccountState();
                        acc.TILOffset += diff;
                        TimeTrackingService.Instance.SetAccountState(acc);
                        var entry = new AccountLogEntry { Date = _activeDay.Date.AddHours(12), Kind = "TIL", Delta = diff, Balance = acc.TILOffset, Note = "Manual override edit", AffectedDate = _activeDay.Date };
                        TimeTrackingService.Instance.AddAccountLogEntry(entry);
                        try { TimeTrackingService.Instance.Reload(); } catch { }
                        RecomputeCumulatives();
                        UpdateAccountsDisplay();
                        UpdateShiftTotalsDisplay();
                        UpdateOverridePanels();
                        DaysList.Items.Refresh();
                        MonthGrid.Items.Refresh();
                    }
                }
            }
            catch { }
        }

        private void DeleteTILOverride_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_activeDay == null) return;
                var current = GetManualOverrideSumForDay(_activeDay.Date.Date, "TIL");
                if (Math.Abs(current) < 0.0001) { UpdateOverridePanels(); return; }
                if (MessageBox.Show("Delete manual override for this day?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
                var acc = TimeTrackingService.Instance.GetAccountState();
                acc.TILOffset -= current;
                TimeTrackingService.Instance.SetAccountState(acc);
                var entry = new AccountLogEntry { Date = _activeDay.Date.AddHours(12), Kind = "TIL", Delta = -current, Balance = acc.TILOffset, Note = "Manual override removed", AffectedDate = _activeDay.Date };
                TimeTrackingService.Instance.AddAccountLogEntry(entry);
                try { TimeTrackingService.Instance.Reload(); } catch { }
                RecomputeCumulatives();
                UpdateAccountsDisplay();
                UpdateShiftTotalsDisplay();
                UpdateOverridePanels();
                DaysList.Items.Refresh();
                MonthGrid.Items.Refresh();
            }
            catch { }
        }

        private void EditHolidayOverride_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_activeDay == null) return;
                var current = GetManualOverrideSumForDay(_activeDay.Date.Date, "Holiday");

                var dlg = new Window
                {
                    Title = "Edit Holiday manual override",
                    SizeToContent = SizeToContent.WidthAndHeight,
                    ResizeMode = ResizeMode.NoResize,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this
                };
                var stack = new StackPanel { Margin = new Thickness(12) };
                stack.Children.Add(new TextBlock { Text = "Enter total manual override for this day:" });
                var inputPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0,8,0,8) };
                var txt = new TextBox { Width = 160, Text = current.ToString("0.##", CultureInfo.CurrentCulture) };
                inputPanel.Children.Add(txt);
                inputPanel.Children.Add(new TextBlock { Text = " days", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6,0,0,0) });
                stack.Children.Add(inputPanel);
                var btns = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
                var ok = new Button { Content = "OK", Width = 80, IsDefault = true, Margin = new Thickness(0,0,6,0) };
                var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
                btns.Children.Add(ok); btns.Children.Add(cancel); stack.Children.Add(btns);
                dlg.Content = stack;

                ok.Click += (s, ea) =>
                {
                    if (double.TryParse(txt.Text, NumberStyles.Any, CultureInfo.CurrentCulture, out var v)) { dlg.Tag = v; dlg.DialogResult = true; }
                    else MessageBox.Show("Enter a valid number", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                };

                if (dlg.ShowDialog() == true)
                {
                    var desired = (double)dlg.Tag;
                    var diff = desired - current;
                    if (Math.Abs(diff) > 0.0001)
                    {
                        var acc = TimeTrackingService.Instance.GetAccountState();
                        acc.HolidayOffset += diff;
                        TimeTrackingService.Instance.SetAccountState(acc);
                        var entry = new AccountLogEntry { Date = _activeDay.Date.AddHours(12), Kind = "Holiday", Delta = diff, Balance = acc.HolidayOffset, Note = "Manual override edit", AffectedDate = _activeDay.Date };
                        TimeTrackingService.Instance.AddAccountLogEntry(entry);
                        try { TimeTrackingService.Instance.Reload(); } catch { }
                        RecomputeCumulatives();
                        UpdateAccountsDisplay();
                        UpdateShiftTotalsDisplay();
                        UpdateOverridePanels();
                        DaysList.Items.Refresh();
                        MonthGrid.Items.Refresh();
                    }
                }
            }
            catch { }
        }

        private void DeleteHolidayOverride_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_activeDay == null) return;
                var current = GetManualOverrideSumForDay(_activeDay.Date.Date, "Holiday");
                if (Math.Abs(current) < 0.0001) { UpdateOverridePanels(); return; }
                if (MessageBox.Show("Delete manual override for this day?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
                var acc = TimeTrackingService.Instance.GetAccountState();
                acc.HolidayOffset -= current;
                TimeTrackingService.Instance.SetAccountState(acc);
                var entry = new AccountLogEntry { Date = _activeDay.Date.AddHours(12), Kind = "Holiday", Delta = -current, Balance = acc.HolidayOffset, Note = "Manual override removed", AffectedDate = _activeDay.Date };
                TimeTrackingService.Instance.AddAccountLogEntry(entry);
                try { TimeTrackingService.Instance.Reload(); } catch { }
                RecomputeCumulatives();
                UpdateAccountsDisplay();
                UpdateShiftTotalsDisplay();
                UpdateOverridePanels();
                DaysList.Items.Refresh();
                MonthGrid.Items.Refresh();
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
                    if (this.FindName("TotalWorkedDataTextBlock") is TextBlock tw) tw.Text = "00:00";
                    if (this.FindName("TotalDeltaDataTextBlock") is TextBlock td) td.Text = "00:00";
                    if (this.FindName("ShiftTargetText") is TextBox st) st.Text = string.Empty;
                    return;
                }

                double totalWorked = 0;
                if (_activeDay.Shifts != null && _activeDay.Shifts.Count > 0)
                    totalWorked = _activeDay.Shifts.Sum(s => s.Hours);
                // no fallback to summary

                double target;
                bool explicitOverride = _activeDay.TargetHours.HasValue;
                if (explicitOverride)
                {
                    target = _activeDay.TargetHours.Value; // can be zero
                }
                else if (_activeDay.DayType == DayType.Vacation || _activeDay.DayType == DayType.PublicHoliday)
                {
                    target = 0; // auto-zero when no explicit override
                }
                else if (_activeDay.Template != null)
                {
                    int idx = ((int)_activeDay.Date.DayOfWeek + 6) % 7;
                    if (_activeDay.Template.HoursPerWeekday != null && _activeDay.Template.HoursPerWeekday.Length == 7)
                        target = _activeDay.Template.HoursPerWeekday[idx];
                    else
                        target = 0;
                }
                else
                {
                    target = 0;
                }

                double delta = totalWorked - (target > 0 ? target : 0);

                if (this.FindName("TotalWorkedDataTextBlock") is TextBlock tw2)
                    tw2.Text = FormatHoursAsHHmm(totalWorked);
                if (this.FindName("TotalDeltaDataTextBlock") is TextBlock td2)
                    td2.Text = FormatHoursAsHHmm(delta);
                if (this.FindName("ShiftTargetText") is TextBox st2)
                {
                    if (explicitOverride)
                        st2.Text = target.ToString("0.##");
                    else
                        st2.Text = target > 0 ? target.ToString("0.##") : "0"; // show 0 when it is auto-zero due to holiday/public holiday
                }
            }
            catch { }
        }

        // New: Export all data to Excel (row per day)
        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var allDays = BuildAllDaysAcrossData();
                if (allDays.Count == 0)
                {
                    MessageBox.Show("No data to export.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var dlg = new SaveFileDialog { Filter = "Excel Workbook (*.xlsx)|*.xlsx", FileName = "TimeTracking.xlsx" };
                if (dlg.ShowDialog(this) != true) return;

                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Days");
                    // Header
                    var headers = new List<string>
                    {
                        "Date","Weekday","Employment Position","Location","Physical Location","Day Type","Target Hours","Worked Hours","Delta","Delta Holiday","Cumulative TIL","Cumulative Holiday",
                        "Manual Reset TIL","Manual Reset Holiday","Manual Reset Notes",
                        "Start1","End1","Lunch1","Desc1",
                        "Start2","End2","Lunch2","Desc2",
                        "Start3","End3","Lunch3","Desc3",
                        "Start4","End4","Lunch4","Desc4"
                    };
                    for (int i = 0; i < headers.Count; i++) ws.Cell(1, i + 1).Value = headers[i];
                    ws.Row(1).Style.Font.Bold = true;
                    ws.SheetView.FreezeRows(1);

                    // cache manual reset log entries by date
                    var allLog = TimeTrackingService.Instance.GetAccountLogEntries(DateTime.MinValue, DateTime.MaxValue, null);

                    int row = 2;
                    foreach (var d in allDays)
                    {
                        // base cells
                        ws.Cell(row, 1).Value = d.Date; ws.Cell(row, 1).Style.DateFormat.Format = "yyyy-MM-dd";
                        ws.Cell(row, 2).Value = d.Date.ToString("dddd");
                        ws.Cell(row, 3).Value = d.PositionOverride ?? d.Template?.JobDescription ?? string.Empty;
                        ws.Cell(row, 4).Value = d.LocationOverride ?? d.Template?.Location ?? string.Empty;
                        ws.Cell(row, 5).Value = d.PhysicalLocationOverride ?? (!string.IsNullOrWhiteSpace(d.LocationOverride) ? d.LocationOverride : d.Template?.Location ?? string.Empty);
                        ws.Cell(row, 6).Value = d.DayType.ToString();
                        ws.Cell(row, 7).Value = d.TargetComputed; ws.Cell(row, 7).Style.NumberFormat.Format = "0.00";
                        ws.Cell(row, 8).Value = d.Worked ?? 0; ws.Cell(row, 8).Style.NumberFormat.Format = "0.00";
                        ws.Cell(row, 9).Value = d.Delta ?? 0; ws.Cell(row, 9).Style.NumberFormat.Format = "0.00";
                        // Delta Holiday: -1 for Vacation, 0 otherwise
                        double deltaHoliday = d.DayType == DayType.Vacation ? -1.0 : 0.0;
                        ws.Cell(row, 10).Value = deltaHoliday; ws.Cell(row, 10).Style.NumberFormat.Format = "0.00";
                        ws.Cell(row, 11).Value = d.CumulativeTIL; ws.Cell(row, 11).Style.NumberFormat.Format = "0.00";
                        ws.Cell(row, 12).Value = d.CumulativeHoliday; ws.Cell(row, 12).Style.NumberFormat.Format = "0.00";

                        // per-day manual resets
                        var dayLogs = allLog.Where(a => a.Date.Date == d.Date.Date && (a.Note ?? string.Empty).IndexOf("manual", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                        double tilReset = dayLogs.Where(a => string.Equals(a.Kind, "TIL", StringComparison.OrdinalIgnoreCase)).Sum(a => a.Delta);
                        double holReset = dayLogs.Where(a => string.Equals(a.Kind, "Holiday", StringComparison.OrdinalIgnoreCase)).Sum(a => a.Delta);
                        string notes = string.Join("; ", dayLogs.Select(a => (a.Kind ?? "") + ": " + a.Delta.ToString("0.##") + (string.IsNullOrWhiteSpace(a.Note) ? "" : " (" + a.Note + ")")));
                        ws.Cell(row, 13).Value = tilReset; ws.Cell(row, 13).Style.NumberFormat.Format = "0.00";
                        ws.Cell(row, 14).Value = holReset; ws.Cell(row, 14).Style.NumberFormat.Format = "0.00";
                        ws.Cell(row, 15).Value = notes;

                        var shifts = (d.Shifts ?? new List<Shift>()).OrderBy(s => s.Start).ToList();
                        for (int i = 0; i < Math.Min(4, shifts.Count); i++)
                        {
                            int c = 16 + i * 4; // Start, End, Lunch, Desc after manual reset columns
                            ws.Cell(row, c + 0).Value = shifts[i].Start.ToString(@"hh\:mm");
                            ws.Cell(row, c + 1).Value = shifts[i].End.ToString(@"hh\:mm");
                            ws.Cell(row, c + 2).Value = shifts[i].LunchBreak.ToString(@"hh\:mm");
                            ws.Cell(row, c + 3).Value = shifts[i].Description ?? string.Empty;
                        }

                        if (d.DayType != DayType.WorkingDay)
                        {
                            var xlColor = DayTypeToXlColor(d.DayType); if (xlColor != null) ws.Row(row).Style.Fill.SetBackgroundColor(xlColor);
                        }
                        row++;
                    }

                    ws.Columns().AdjustToContents();

                    var logWs = wb.Worksheets.Add("ManualResets");
                    logWs.Cell(1, 1).Value = "Date"; logWs.Cell(1, 2).Value = "Kind"; logWs.Cell(1, 3).Value = "Delta"; logWs.Cell(1, 4).Value = "Balance"; logWs.Cell(1, 5).Value = "Note"; logWs.Cell(1, 6).Value = "AffectedDate"; logWs.Row(1).Style.Font.Bold = true;
                    int lr = 2; foreach (var a in allLog.Where(a => (a.Note ?? string.Empty).IndexOf("manual", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        logWs.Cell(lr, 1).Value = a.Date; logWs.Cell(lr, 1).Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
                        logWs.Cell(lr, 2).Value = a.Kind; logWs.Cell(lr, 3).Value = a.Delta; logWs.Cell(lr, 3).Style.NumberFormat.Format = "0.00";
                        logWs.Cell(lr, 4).Value = a.Balance; logWs.Cell(lr, 4).Style.NumberFormat.Format = "0.00";
                        logWs.Cell(lr, 5).Value = a.Note ?? string.Empty; logWs.Cell(lr, 6).Value = a.AffectedDate.HasValue ? a.AffectedDate.Value.ToString("yyyy-MM-dd") : string.Empty; lr++;
                    }
                    logWs.Columns().AdjustToContents();

                    wb.SaveAs(dlg.FileName);
                }
                MessageBox.Show("Exported to Excel.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex) { MessageBox.Show("Failed to export: " + ex.Message, "Export", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private void ImportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new OpenFileDialog { Filter = "Excel Workbook (*.xlsx)|*.xlsx" }; if (dlg.ShowDialog(this) != true) return;
                var pending = new List<(DateTime Date, string Kind, double Delta, string Note)>();
                using (var wb = new XLWorkbook(dlg.FileName))
                {
                    var ws = wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, "Days", StringComparison.OrdinalIgnoreCase)) ?? wb.Worksheet(1);
                    if (ws == null) { MessageBox.Show("Worksheet not found.", "Import", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
                    var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase); var used = ws.RangeUsed(); if (used == null) { MessageBox.Show("No data in sheet.", "Import", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
                    var firstRow = used.FirstRowUsed(); int lastCol = firstRow.LastCellUsed().Address.ColumnNumber; for (int c = 1; c <= lastCol; c++) { var name = (firstRow.Cell(c).GetString() ?? string.Empty).Trim(); if (!string.IsNullOrEmpty(name) && !headerMap.ContainsKey(name)) headerMap.Add(name, c); }
                    int Col(string name) { return headerMap.TryGetValue(name, out var cc) ? cc : 0; }
                    int dateCol = Col("Date"); if (dateCol == 0) { MessageBox.Show("Required column 'Date' not found.", "Import", MessageBoxButton.OK, MessageBoxImage.Warning); return; }

                    for (int r = firstRow.RowBelow().RowNumber(); r <= used.LastRowUsed().RowNumber(); r++)
                    {
                        var cell = ws.Cell(r, dateCol); if (cell == null) continue; DateTime date; if (cell.TryGetValue<DateTime>(out var dtVal)) date = dtVal.Date; else { var s = cell.GetString(); if (!DateTime.TryParse(s, out date)) continue; date = date.Date; }
                        string pos = GetString(ws, r, Col("Employment Position")); string loc = GetString(ws, r, Col("Location")); string phys = GetString(ws, r, Col("Physical Location")); string dayTypeStr = GetString(ws, r, Col("Day Type")); DayType? dayType = TryParseDayType(dayTypeStr); double? target = GetDouble(ws, r, Col("Target Hours"));
                        var shifts = new List<Shift>();
                        for (int i = 1; i <= 10; i++)
                        {
                            int baseCol = Col($"Start{i}"); if (baseCol == 0) break;
                            var startStr = GetString(ws, r, baseCol); var endStr = GetString(ws, r, baseCol + 1); var lunchStr = GetString(ws, r, baseCol + 2); var desc = GetString(ws, r, baseCol + 3);
                            if (string.IsNullOrWhiteSpace(startStr) && string.IsNullOrWhiteSpace(endStr) && string.IsNullOrWhiteSpace(lunchStr) && string.IsNullOrWhiteSpace(desc)) continue;
                            if (!TimeSpan.TryParse(startStr, out var startTs)) continue; if (!TimeSpan.TryParse(endStr, out var endTs)) continue; if (!TimeSpan.TryParse(string.IsNullOrWhiteSpace(lunchStr) ? "00:00" : lunchStr, out var lunchTs)) lunchTs = TimeSpan.Zero;
                            shifts.Add(new Shift { Date = date, Start = startTs, End = endTs, LunchBreak = lunchTs, Description = desc ?? string.Empty, DayMode = "import", ManualStartOverride = false, ManualEndOverride = false });
                        }
                        TimeTrackingService.Instance.UpsertOverride(date, pos ?? string.Empty, loc ?? string.Empty, phys ?? (loc ?? string.Empty), target, dayType);
                        TimeTrackingService.Instance.UpsertShiftsForDate(date, shifts);

                        // read per-day manual reset columns (optional)
                        var tilReset = GetDouble(ws, r, Col("Manual Reset TIL"));
                        var holReset = GetDouble(ws, r, Col("Manual Reset Holiday"));
                        var note = GetString(ws, r, Col("Manual Reset Notes"));
                        if (tilReset.HasValue && Math.Abs(tilReset.Value) > 0.0001) pending.Add((date.AddHours(12), "TIL", tilReset.Value, string.IsNullOrWhiteSpace(note) ? "Imported manual reset" : note));
                        if (holReset.HasValue && Math.Abs(holReset.Value) > 0.0001) pending.Add((date.AddHours(12), "Holiday", holReset.Value, string.IsNullOrWhiteSpace(note) ? "Imported manual reset" : note));
                    }
                }

                // Create account log entries for pending manual resets with computed running balances
                if (pending.Count > 0)
                {
                    var existing = TimeTrackingService.Instance.GetAccountLogEntries(DateTime.MinValue, DateTime.MaxValue, null).ToList();
                    var byKind = pending.GroupBy(p => p.Kind, StringComparer.OrdinalIgnoreCase);
                    foreach (var group in byKind)
                    {
                        string kind = group.Key;
                        var items = group.OrderBy(p => p.Date).ToList();
                        var firstDate = items.First().Date.Date;
                        var savedBefore = TimeTrackingService.Instance.GetAccountLogEntries(DateTime.MinValue, firstDate.AddDays(-1), kind).OrderBy(a => a.Date).ToList();
                        double start = savedBefore.Count > 0 ? savedBefore.Last().Balance : 0.0;
                        double running = start;
                        foreach (var p in items)
                        {
                            // avoid duplicate: if an existing manual entry on same date with same delta and kind exists, skip
                            if (existing.Any(e => e.Date.Date == p.Date.Date && e.Kind.Equals(kind, StringComparison.OrdinalIgnoreCase) && Math.Abs(e.Delta - p.Delta) < 0.0001 && (e.Note ?? "").IndexOf("manual", StringComparison.OrdinalIgnoreCase) >= 0))
                                continue;
                            running += p.Delta;
                            var entry = new AccountLogEntry { Date = p.Date, Kind = kind, Delta = p.Delta, Balance = running, Note = p.Note, AffectedDate = p.Date.Date };
                            TimeTrackingService.Instance.AddAccountLogEntry(entry);
                        }
                    }
                }

                try { TimeTrackingService.Instance.Reload(); } catch { }
                PopulateDaysForCurrentMonth();
                RecomputeCumulatives();
                UpdateAccountsDisplay();
                UpdateShiftTotalsDisplay();
                UpdateOverridePanels();
                MessageBox.Show("Import completed.", "Import", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to import: " + ex.Message, "Import", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<DayViewModel> BuildAllDaysAcrossData()
        {
            var days = new List<DayViewModel>();
            try
            {
                DateTime? min = null; DateTime? max = null;
                try { var events = TimeTrackingService.Instance.GetEvents(); if (events != null && events.Count > 0) { var mn = events.Min(e => e.Timestamp.Date); var mx = events.Max(e => e.Timestamp.Date); min = min.HasValue ? (mn < min ? mn : min) : mn; max = max.HasValue ? (mx > max ? mx : max) : mx; } } catch { }
                try { var sh = TimeTrackingService.Instance.GetAllSavedShifts(); if (sh != null && sh.Count > 0) { var mn = sh.Min(s => s.Date.Date); var mx = sh.Max(s => s.Date.Date); min = min.HasValue ? (mn < min ? mn : min) : mn; max = max.HasValue ? (mx > max ? mx : max) : mx; } } catch { }
                try { var ov = TimeTrackingService.Instance.GetOverrides(); if (ov != null && ov.Count > 0) { var mn = ov.Min(o => o.Date.Date); var mx = ov.Max(o => o.Date.Date); min = min.HasValue ? (mn < min ? mn : min) : mn; max = max.HasValue ? (mx > max ? mx : max) : mx; } } catch { }
                try { var tmps = TimeTrackingService.Instance.GetTemplates(); if (tmps != null && tmps.Count > 0) { var mn = tmps.Min(t => t.StartDate.Date); var mx = tmps.Max(t => (t.EndDate ?? DateTime.Today).Date); min = min.HasValue ? (mn < min ? mn : min) : mn; max = max.HasValue ? (mx > max ? mx : max) : mx; } } catch { }
                if (!min.HasValue || !max.HasValue)
                {
                    var f = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1); var l = f.AddMonths(1).AddDays(-1); var summaries0 = TimeTrackingService.Instance.GetDaySummaries(f, l); days = summaries0.Select(s => new DayViewModel(s)).ToList();
                }
                else { var summaries = TimeTrackingService.Instance.GetDaySummaries(min.Value, max.Value); days = summaries.Select(s => new DayViewModel(s)).ToList(); }
                double runningTIL = TimeTrackingService.Instance.GetAccountState().TILOffset; double runningHoliday = TimeTrackingService.Instance.GetAccountState().HolidayOffset; var ordered = days.OrderBy(d => d.Date).ToList();
                foreach (var d in ordered)
                {
                    var ov = TimeTrackingService.Instance.GetOverrideForDate(d.Date);
                    if (ov != null)
                    {
                        d.PositionOverride = ov.Position; d.LocationOverride = ov.Location; d.PhysicalLocationOverride = string.IsNullOrWhiteSpace(ov.PhysicalLocation) ? (string.IsNullOrWhiteSpace(ov.Location) ? d.Template?.Location : ov.Location) : ov.PhysicalLocation; d.TargetHours = ov.TargetHours; if (ov.DayType.HasValue) d.SetDayType(ov.DayType.Value);
                    }
                    else if (d.Template != null && !string.IsNullOrWhiteSpace(d.Template.Location)) d.PhysicalLocationOverride = d.Template.Location;
                    var shs = TimeTrackingService.Instance.GetShiftsForDate(d.Date).ToList(); d.Shifts = shs.Any() ? shs.Select(s2 => new Shift { Date = s2.Date, Start = s2.Start, End = s2.End, Description = s2.Description, DayMode = s2.DayMode, ManualStartOverride = s2.ManualStartOverride, ManualEndOverride = s2.ManualEndOverride, LunchBreak = s2.LunchBreak }).ToList() : new List<Shift>();

                    // For import, always reset cumulatives to avoid multiplication effect from repeated imports
                    d.CumulativeTIL = 0;
                    d.CumulativeHoliday = 0;

                    double worked = d.Shifts.Sum(s => s.Hours); double target; if (d.TargetHours.HasValue) target = d.TargetHours.Value; else if (d.DayType == DayType.Vacation || d.DayType == DayType.PublicHoliday) target = 0; else if (d.Template != null && d.Template.HoursPerWeekday?.Length == 7) { int idx = ((int)d.Date.DayOfWeek + 6) % 7; target = d.Template.HoursPerWeekday[idx]; } else target = 0; double delta = worked - (target > 0 ? target : 0);
                    if (d.DayType == DayType.WorkingDay) runningTIL += delta; else if (d.DayType == DayType.Vacation) runningHoliday -= 1.0; d.CumulativeTIL = runningTIL; d.CumulativeHoliday = runningHoliday;
                }
                days = ordered;
            }
            catch { }
            return days;
        }

        private XLColor? DayTypeToXlColor(DayType dt)
        {
            switch (dt)
            {
                case DayType.Weekend: return XLColor.FromHtml("#F0F0F0");
                case DayType.PublicHoliday: return XLColor.FromHtml("#FFE5E5");
                case DayType.Vacation: return XLColor.FromHtml("#DDEEFF");
                case DayType.TimeInLieu: return XLColor.FromHtml("#FFF2CC");
                case DayType.Other: return XLColor.FromHtml("#EFEFEF");
                default: return null;
            }
        }

        private static DayType? TryParseDayType(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            s = s.Trim();
            // allow spaces-insensitive match
            string norm = new string(s.Where(char.IsLetter).ToArray());
            foreach (DayType dt in Enum.GetValues(typeof(DayType)))
            {
                string n = new string(dt.ToString().Where(char.IsLetter).ToArray());
                if (string.Equals(n, norm, StringComparison.OrdinalIgnoreCase)) return dt;
            }
            return null;
        }

        private static string GetString(IXLWorksheet ws, int row, int col) { if (col <= 0) return null; var c = ws.Cell(row, col); return c?.GetString(); }
        private static double? GetDouble(IXLWorksheet ws, int row, int col)
        {
            if (col <= 0) return null; var c = ws.Cell(row, col); if (c.TryGetValue<double>(out var d)) return d; var s = c.GetString(); if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d)) return d; if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return d; return null;
        }

        // Format hours as HH:MM for UI/account labels
        public static string FormatHoursAsHHmm(double hours)
        {
            if (double.IsNaN(hours) || double.IsInfinity(hours)) return "00:00";
            var negative = hours < 0; var abs = Math.Abs(hours);
            int h = (int)Math.Floor(abs);
            int m = (int)Math.Round((abs - h) * 60);
            if (m == 60) { h += 1; m = 0; }
            var s = $"{h:00}:{m:00}";
            return negative ? "-" + s : s;
        }

        // Refresh account values shown in the UI (Holiday and TIL)
        private void UpdateAccountsDisplay()
        {
            try
            {
                double holidayVal, tilVal;
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
                if (this.FindName("AccountHolidayValue") is TextBlock ah) ah.Text = holidayVal.ToString("0.##") + " days";
                if (this.FindName("AccountTILValue") is TextBlock til) til.Text = FormatHoursAsHHmm(tilVal);
                DaysList.Items.Refresh();
                MonthGrid.Items.Refresh();
            }
            catch { }
        }

        // Toggle handlers for colouring by physical location
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

        // Export a PDF with year maps (day type, physical location, and overtime)
        private void ExportYearMap_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Build available years based on data
                var years = GetAvailableYears().ToList();
                if (years.Count == 0)
                {
                    MessageBox.Show("No data to export.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Preselect current year if present
                IEnumerable<int> pre = Enumerable.Empty<int>();
                if (YearCombo.SelectedItem is string ys && int.TryParse(ys, out var curYear) && years.Contains(curYear)) pre = new[] { curYear };

                var picker = new YearSelectionDialog(years, pre) { Owner = this };
                if (picker.ShowDialog() != true) return;
                var selectedYears = picker.SelectedYears.ToList();
                if (selectedYears.Count == 0) return;

                // Ask where to save
                string defaultName = selectedYears.Count == 1 ? $"timemap_{selectedYears[0]}.pdf" : $"timemap_{selectedYears.First()}-{selectedYears.Last()}.pdf";
                var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "PDF files (*.pdf)|*.pdf", FileName = defaultName };
                if (sfd.ShowDialog() != true) return;

                using (var doc = new PdfDocument())
                {
                    // Group pages: first Day Type for all years, then Location for all years, then Hours for all years
                    // Day Type pages
                    foreach (var y in selectedYears)
                    {
                        var days = BuildYearDays(y);
                        var p = doc.AddPage(); p.Size = PdfSharpCore.PageSize.A4; using (var g = XGraphics.FromPdfPage(p)) { DrawYearGridToGraphics(g, days, y, colorByPhysical: false); }
                    }
                    // Location pages
                    foreach (var y in selectedYears)
                    {
                        var days = BuildYearDays(y);
                        var p = doc.AddPage(); p.Size = PdfSharpCore.PageSize.A4; using (var g = XGraphics.FromPdfPage(p)) { DrawYearGridToGraphics(g, days, y, colorByPhysical: true); }
                    }
                    // Hours pages
                    foreach (var y in selectedYears)
                    {
                        var days = BuildYearDays(y);
                        var p = doc.AddPage(); p.Size = PdfSharpCore.PageSize.A4; using (var g = XGraphics.FromPdfPage(p)) { DrawOvertimeGridToGraphics(g, days, y); }
                    }

                    using (var fs = File.OpenWrite(sfd.FileName)) { doc.Save(fs); }
                }
                MessageBox.Show("Exported colour maps to PDF.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to export: " + ex.Message, "Export", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Helpers for PDF year export
        private List<DayViewModel> BuildYearDays(int year)
        {
            var first = new DateTime(year, 1, 1);
            var last = new DateTime(year, 12, 31);
            var summaries = TimeTrackingService.Instance.GetDaySummaries(first, last);
            var days = summaries.Select(s => new DayViewModel(s)).ToList();
            foreach (var d in days)
            {
                var ov = TimeTrackingService.Instance.GetOverrideForDate(d.Date);
                if (ov != null)
                {
                    if (!string.IsNullOrWhiteSpace(ov.PhysicalLocation)) d.PhysicalLocationOverride = ov.PhysicalLocation;
                    else if (!string.IsNullOrWhiteSpace(ov.Location)) d.PhysicalLocationOverride = ov.Location;
                    d.PositionOverride = ov.Position; d.LocationOverride = ov.Location; d.TargetHours = ov.TargetHours; if (ov.DayType.HasValue) d.SetDayType(ov.DayType.Value);
                }
                var saved = TimeTrackingService.Instance.GetShiftsForDate(d.Date).ToList();
                d.Shifts = saved.Any() ? saved.Select(s2 => new Shift { Date = s2.Date, Start = s2.Start, End = s2.End, Description = s2.Description, ManualStartOverride = s2.ManualStartOverride, ManualEndOverride = s2.ManualEndOverride, LunchBreak = s2.LunchBreak, DayMode = s2.DayMode }).ToList() : new List<Shift>();
            }
            return days;
        }

        private IEnumerable<int> GetAvailableYears()
        {
            var years = new HashSet<int>();
            try
            {
                var events = TimeTrackingService.Instance.GetEvents();
                if (events != null) foreach (var e in events) years.Add(e.Timestamp.Year);
            }
            catch { }
            try
            {
                var sh = TimeTrackingService.Instance.GetAllSavedShifts();
                if (sh != null) foreach (var s in sh) years.Add(s.Date.Year);
            }
            catch { }
            try
            {
                var ov = TimeTrackingService.Instance.GetOverrides();
                if (ov != null) foreach (var o in ov) years.Add(o.Date.Year);
            }
            catch { }
            try
            {
                var tmps = TimeTrackingService.Instance.GetTemplates();
                if (tmps != null)
                {
                    foreach (var t in tmps)
                    {
                        var startY = t.StartDate.Year; var endY = (t.EndDate ?? DateTime.Today).Year;
                        for (int y = startY; y <= endY; y++) years.Add(y);
                    }
                }
            }
            catch { }

            // Also include the current year selection and today for convenience
            try { years.Add(DateTime.Today.Year); } catch { }
            if (YearCombo.SelectedItem is string ys && int.TryParse(ys, out var comboYear)) years.Add(comboYear);

            return years.OrderBy(y => y);
        }

        private void DrawYearGridToGraphics(XGraphics g, List<DayViewModel> days, int year, bool colorByPhysical)
        {
            double margin = 30; double pageWidth = g.PageSize.Width; double pageHeight = g.PageSize.Height; double contentWidth = pageWidth - (margin * 2); double legendHeight = 70;
            var headerFont = new XFont("Arial", 18, XFontStyle.Bold); g.DrawString((colorByPhysical ? "Physical location map - " : "Day type map - ") + year.ToString(), headerFont, XBrushes.Black, new XRect(margin, 10, contentWidth, 30), XStringFormats.TopLeft);
            int cols = 3; int rows = 4; double gap = 8; double monthWidth = (contentWidth - (gap * (cols - 1))) / cols; double monthGridAvailableHeight = pageHeight - 80 - legendHeight - (gap * (rows - 1)); double monthHeight = monthGridAvailableHeight / rows;
            for (int m = 1; m <= 12; m++) { int monthIndex = m - 1; int col = monthIndex % cols; int row = monthIndex / cols; double x = margin + col * (monthWidth + gap); double y = 50 + row * (monthHeight + gap); DrawSingleMonth(g, days, year, m, new XRect(x, y, monthWidth, monthHeight), colorByPhysical); }
            DrawLegend(g, margin, pageHeight - legendHeight + 12, colorByPhysical, days, contentWidth);
        }

        private void DrawSingleMonth(XGraphics g, List<DayViewModel> days, int year, int month, XRect rect, bool colorByPhysical)
        {
            var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(month); var titleFont = new XFont("Arial", 12, XFontStyle.Bold); g.DrawString(monthName + " " + year.ToString(), titleFont, XBrushes.Black, new XPoint(rect.X + 4, rect.Y + 14));
            double top = rect.Y + 22; double left = rect.X + 2; double gridWidth = rect.Width - 4; double gridHeight = rect.Height - 28; int cols = 7; int rows = 6; double cellW = gridWidth / cols; double cellH = gridHeight / rows; var first = new DateTime(year, month, 1); int leading = (int)first.DayOfWeek - 1; if (leading < 0) leading += 7;
            for (int i = 0; i < rows * cols; i++)
            {
                int dayNum = i - leading + 1; double cx = left + (i % cols) * cellW; double cy = top + (i / cols) * cellH; var cellRect = new XRect(cx, cy, cellW - 1, cellH - 1);
                if (dayNum >= 1 && dayNum <= DateTime.DaysInMonth(year, month))
                {
                    var d = days.FirstOrDefault(dd => dd.Date.Year == year && dd.Date.Month == month && dd.Date.Day == dayNum);
                    XBrush brush = XBrushes.White;
                    if (d != null)
                    {
                        if (colorByPhysical)
                        {
                            var key = !string.IsNullOrWhiteSpace(d.PhysicalLocationOverride) ? d.PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(d.LocationOverride) ? d.LocationOverride : d.Template?.Location ?? "");
                            if (string.IsNullOrWhiteSpace(key)) brush = XBrushes.White; else { var hex = TimeTrackingService.Instance.GetLocationColor(key); if (!string.IsNullOrWhiteSpace(hex)) { try { var sysc = (Color)ColorConverter.ConvertFromString(hex); brush = new XSolidBrush(XColor.FromArgb(sysc.A, sysc.R, sysc.G, sysc.B)); } catch { brush = XBrushes.White; } } else { int h = key.GetHashCode(); byte r = (byte)(80 + (Math.Abs(h) % 176)); byte gcol = (byte)(80 + (Math.Abs(h / 7) % 176)); byte b = (byte)(80 + (Math.Abs(h / 13) % 176)); brush = new XSolidBrush(XColor.FromArgb(255, r, gcol, b)); } }
                        }
                        else { if (DayTypePdfColors.TryGetValue(d.DayType, out var c)) brush = new XSolidBrush(c); else brush = XBrushes.White; }
                    }
                    g.DrawRectangle(brush, cellRect); var numFont = new XFont("Arial", 8, XFontStyle.Regular); g.DrawString(dayNum.ToString(), numFont, XBrushes.Black, new XPoint(cellRect.X + 4, cellRect.Y + 10));
                }
                else g.DrawRectangle(XBrushes.LightGray, cellRect);
            }
            g.DrawRectangle(XPens.Black, rect.X, rect.Y, rect.Width, rect.Height);
        }

        private void DrawLegend(XGraphics g, double x, double y, bool colorByPhysical, List<DayViewModel> days, double availableWidth)
        {
            var font = new XFont("Arial", 9, XFontStyle.Regular); double box = 12; double gap = 8; double cx = x; double startX = x; double padding = 6;
            if (!colorByPhysical)
            {
                var items = DayTypePdfColors.Keys.Select(k => k.ToString()).ToList(); var widths = items.Select(it => g.MeasureString(it, font).Width + box + gap + padding).ToList(); double totalW = widths.Sum(); cx = startX + Math.Max(0, (availableWidth - totalW) / 2);
                for (int i = 0; i < items.Count; i++) { var key = items[i]; var brush = new XSolidBrush(DayTypePdfColors[(DayType)Enum.Parse(typeof(DayType), key)]); g.DrawRectangle(brush, cx, y, box, box); g.DrawString(key, font, XBrushes.Black, new XPoint(cx + box + 6, y + box - 2)); cx += widths[i]; }
            }
            else
            {
                var locs = daysForLegendLookup(); int max = Math.Min(8, locs.Count); var widths = new List<double>(); for (int i = 0; i < max; i++) widths.Add(g.MeasureString(locs[i], font).Width + box + gap + padding); double totalW = widths.Sum(); cx = startX + Math.Max(0, (availableWidth - totalW) / 2);
                for (int i = 0; i < max; i++) { var key = locs[i]; var hex = TimeTrackingService.Instance.GetLocationColor(key); XBrush brush; if (!string.IsNullOrWhiteSpace(hex)) { try { var sysc = (Color)ColorConverter.ConvertFromString(hex); brush = new XSolidBrush(XColor.FromArgb(sysc.A, sysc.R, sysc.G, sysc.B)); } catch { brush = XBrushes.White; } } else { int h = key.GetHashCode(); byte r = (byte)(80 + (Math.Abs(h) % 176)); byte gcol = (byte)(80 + (Math.Abs(h / 7) % 176)); byte b = (byte)(80 + (Math.Abs(h / 13) % 176)); brush = new XSolidBrush(XColor.FromArgb(255, r, gcol, b)); } g.DrawRectangle(brush, cx, y, box, box); g.DrawString(key, font, XBrushes.Black, new XPoint(cx + box + 6, y + box - 2)); cx += widths[i]; }
            }
            List<string> daysForLegendLookup() { var locs = days.Where(d => !string.IsNullOrWhiteSpace(d.PhysicalLocationOverride) || !string.IsNullOrWhiteSpace(d.LocationOverride) || !string.IsNullOrWhiteSpace(d.Template?.Location)).Select(d => !string.IsNullOrWhiteSpace(d.PhysicalLocationOverride) ? d.PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(d.LocationOverride) ? d.LocationOverride : d.Template?.Location ?? "")).Where(s => !string.IsNullOrWhiteSpace(s)).GroupBy(s => s).OrderByDescending(gp => gp.Count()).Select(gp => gp.Key).ToList(); return locs; }
        }

        private void DrawOvertimeGridToGraphics(XGraphics g, List<DayViewModel> days, int year)
        {
            double margin = 30; double pageWidth = g.PageSize.Width; double pageHeight = g.PageSize.Height; double contentWidth = pageWidth - (margin * 2);
            var headerFont = new XFont("Arial", 16, XFontStyle.Bold); var subFont = new XFont("Arial", 10, XFontStyle.Regular);
            double min = -8.0; double max = 8.0; double totalOver = days.Where(d => d.Delta.HasValue && d.Delta.Value > 0).Sum(d => d.Delta.Value); int countOver = days.Count(d => d.Delta.HasValue && d.Delta.Value > 0);
            g.DrawString("Overtime (Delta) map - " + year.ToString(), headerFont, XBrushes.Black, new XRect(margin, 10, contentWidth, 24), XStringFormats.TopLeft);
            g.DrawString($"Total overtime: {totalOver:0.##}h across {countOver} day(s)", subFont, XBrushes.Black, new XRect(margin, 34, contentWidth, 20), XStringFormats.TopLeft);
            int cols = 3; int rows = 4; double gap = 8; double monthWidth = (contentWidth - (gap * (cols - 1))) / cols; double legendHeight = 70; double monthGridAvailableHeight = pageHeight - 80 - legendHeight - (gap * (rows - 1)); double monthHeight = monthGridAvailableHeight / rows;
            for (int m = 1; m <= 12; m++) { int monthIndex = m - 1; int col = monthIndex % cols; int row = monthIndex / cols; double x = margin + col * (monthWidth + gap); double y = 60 + row * (monthHeight + gap); DrawSingleMonthOvertime(g, days, year, m, new XRect(x, y, monthWidth, monthHeight), min, max); }
            double legendY = pageHeight - 60; double legendX = margin + 20; double legendW = contentWidth - 40; double legendH = 14; int steps = Math.Max(4, (int)legendW); for (int i = 0; i < steps; i++) { double t = (double)i / Math.Max(1, steps - 1); double val = min + t * (max - min); var c = JetColor(val, min, max); var brush = new XSolidBrush(c); g.DrawRectangle(brush, legendX + i, legendY, 1, legendH); }
            var lblFont = new XFont("Arial", 9, XFontStyle.Regular); g.DrawString(min.ToString("0.##") + "h", lblFont, XBrushes.Black, new XPoint(legendX, legendY + legendH + 6)); double zeroPos = legendX; if (max > min) zeroPos = legendX + legendW * ((0 - min) / (max - min)); g.DrawString("0h", lblFont, XBrushes.Black, new XPoint(zeroPos - g.MeasureString("0h", lblFont).Width/2, legendY + legendH + 6)); g.DrawString(max.ToString("0.##") + "h", lblFont, XBrushes.Black, new XPoint(legendX + legendW - g.MeasureString(max.ToString("0.##") + "h", lblFont).Width, legendY + legendH + 6));
        }

        private void DrawSingleMonthOvertime(XGraphics g, List<DayViewModel> days, int year, int month, XRect rect, double min, double max)
        {
            var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(month); var titleFont = new XFont("Arial", 10, XFontStyle.Bold); g.DrawString(monthName + " " + year.ToString(), titleFont, XBrushes.Black, new XPoint(rect.X + 4, rect.Y + 12));
            double top = rect.Y + 20; double left = rect.X + 2; double gridWidth = rect.Width - 4; double gridHeight = rect.Height - 24; int cols = 7; int rows = 6; double cellW = gridWidth / cols; double cellH = gridHeight / rows; var first = new DateTime(year, month, 1); int leading = (int)first.DayOfWeek - 1; if (leading < 0) leading += 7;
            for (int i = 0; i < rows * cols; i++)
            {
                int dayNum = i - leading + 1; double cx = left + (i % cols) * cellW; double cy = top + (i / cols) * cellH; var cellRect = new XRect(cx, cy, cellW - 1, cellH - 1);
                if (dayNum >= 1 && dayNum <= DateTime.DaysInMonth(year, month))
                {
                    var d = days.FirstOrDefault(dd => dd.Date.Year == year && dd.Date.Month == month && dd.Date.Day == dayNum);
                    if (d != null && d.Delta.HasValue) { var jc = JetColor(d.Delta.Value, min, max); g.DrawRectangle(new XSolidBrush(jc), cellRect); }
                    else if (d != null) g.DrawRectangle(XBrushes.White, cellRect); else g.DrawRectangle(XBrushes.LightGray, cellRect);
                    var numFont = new XFont("Arial", 7, XFontStyle.Regular); g.DrawString(dayNum.ToString(), numFont, XBrushes.Black, new XPoint(cellRect.X + 3, cellRect.Y + 9));
                }
                else g.DrawRectangle(XBrushes.LightGray, cellRect);
            }
            g.DrawRectangle(XPens.Black, rect.X, rect.Y, rect.Width, rect.Height);
        }

        private XColor JetColor(double value, double min, double max)
        {
            if (value < min || value > max) return XColor.FromArgb(255, 0, 0, 0);
            double t = (max > min) ? (value - min) / (max - min) : 0.5; t = Math.Min(1, Math.Max(0, t));
            double r = 0, g = 0, b = 0; if (t < 0.25) { r = 0; g = 4 * t; b = 1; } else if (t < 0.5) { r = 0; g = 1; b = 1 - 4 * (t - 0.25); } else if (t < 0.75) { r = 4 * (t - 0.5); g = 1; b = 0; } else { r = 1; g = 1 - 4 * (t - 0.75); b = 0; }
            byte R = (byte)(Math.Min(1, Math.Max(0, r)) * 255); byte G = (byte)(Math.Min(1, Math.Max(0, g)) * 255); byte B = (byte)(Math.Min(1, Math.Max(0, b)) * 255); return XColor.FromArgb(255, R, G, B);
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
        public double? Worked { get { try { if (Shifts != null && Shifts.Count > 0) return Shifts.Sum(s => s.Hours); } catch { } return 0; } }
        public double Standard { get; }
        public double CumulativeTIL { get; set; }
        public double CumulativeHoliday { get; set; }
        public string CumulativeTILDisplay => TimeTrackingDialog.FormatHoursAsHHmm(CumulativeTIL);
        public string CumulativeHolidayDisplay => CumulativeHoliday.ToString("0.##") + "d";
        public double? Delta => Worked.HasValue ? (Worked.Value - TargetComputed) : (double?)null;
        public string WorkedDisplay => Worked.HasValue ? TimeTrackingDialog.FormatHoursAsHHmm(Worked.Value) : "00:00";
        public string StandardDisplay => Standard.ToString("0.00");
        public DayType DayType { get; private set; }
        public TimeTemplate Template { get; }
        public bool IsToday { get; set; }
        public bool IsSelected { get; set; }
        public string PositionOverride { get; set; }
        public string LocationOverride { get; set; }
        public string PhysicalLocationOverride { get; set; }
        public List<Shift> Shifts { get; set; }
        public double? TargetHours { get; set; }
        public double TargetComputed { get { if (TargetHours.HasValue) return TargetHours.Value; if (DayType == DayType.Vacation || DayType == DayType.PublicHoliday) return 0; if (Template != null && Template.HoursPerWeekday != null && Template.HoursPerWeekday.Length == 7) { int idx = ((int)Date.DayOfWeek + 6) % 7; return Template.HoursPerWeekday[idx]; } return 0; } }
        public string TargetDisplay => TargetComputed.ToString("0.00");

        public DayViewModel(DaySummary s)
        {
            Date = s.Date; Standard = s.StandardHours; Template = TimeTrackingService.Instance.GetTemplates().FirstOrDefault(t => t.AppliesTo(s.Date)); DayType = (Date.DayOfWeek == DayOfWeek.Saturday || Date.DayOfWeek == DayOfWeek.Sunday) ? DayType.Weekend : DayType.WorkingDay; IsToday = false; IsSelected = false; PositionOverride = null; LocationOverride = null; PhysicalLocationOverride = null;
            var saved = TimeTrackingService.Instance.GetShiftsForDate(Date).ToList(); Shifts = (saved != null && saved.Count > 0) ? saved.Select(s2 => new Shift { Date = s2.Date, Start = s2.Start, End = s2.End, Description = s2.Description, DayMode = s2.DayMode, ManualStartOverride = s2.ManualStartOverride, ManualEndOverride = s2.ManualEndOverride, LunchBreak = s2.LunchBreak }).ToList() : new List<Shift>();
        }

        public void SetDayType(DayType dt) => DayType = dt;

        public Brush BackgroundBrush
        {
            get
            {
                try
                {
                    if (TimeTrackingDialog.ColorByPhysicalLocation)
                    {
                        var key = !string.IsNullOrWhiteSpace(PhysicalLocationOverride) ? PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(LocationOverride) ? LocationOverride : Template?.Location ?? "");
                        if (string.IsNullOrWhiteSpace(key)) return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFFFFF"));

                        // use saved color if present
                        var hex = TimeTrackingService.Instance.GetLocationColor(key);
                        if (!string.IsNullOrWhiteSpace(hex)) { try { var c = (Color)ColorConverter.ConvertFromString(hex); return new SolidColorBrush(c); } catch { } }

                        int h = key.GetHashCode(); byte r = (byte)(80 + (Math.Abs(h) % 176)); byte gcol = (byte)(80 + (Math.Abs(h / 7) % 176)); byte b = (byte)(80 + (Math.Abs(h / 13) % 176)); return new SolidColorBrush(Color.FromRgb(r, gcol, b));
                    }
                    else
                    {
                        switch (DayType)
                        {
                            case DayType.WorkingDay: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFFFFF"));
                            case DayType.Weekend: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF0F0F0"));
                            case DayType.PublicHoliday: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFE5E5"));
                            case DayType.Vacation: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFDDEEFF"));
                            case DayType.TimeInLieu: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFF2CC"));
                            case DayType.Other: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFEFEFEF"));
                            default: return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFFFFF"));
                        }
                    }
                }
                catch { return new SolidColorBrush(Colors.White); }
            }
        }
    }
 }

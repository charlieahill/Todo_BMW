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
        private static bool _holidayBackfillDone = false; // no longer used, kept for compatibility

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

        // Export map modes
        private enum ExportMapMode
        {
            DayType,
            PhysicalLocation,
            OvertimeGradient
        }

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

            // Load persisted overrides and shifts, compute per-day worked/target, but do not compute cumulatives here
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

                // Initialize cumulatives to 0; they will be recomputed below
                d.CumulativeTIL = 0;
                d.CumulativeHoliday = 0;
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

            // Compute cumulative balances using new reset-aware logic across all data, then map in-month
            RecomputeCumulatives();
            UpdateAccountsDisplay();
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

                        // NOTE: We no longer create or depend on generated holiday log entries.
                        // Holiday balance is computed directly from Vacation day markings during recalculation.

                        DaysList.Items.Refresh();
                        MonthGrid.Items.Refresh();
                        // Recompute cumulatives and update accounts since day types/targets changed
                        RecomputeCumulativesFrom(vm.Date);
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

                // Recompute cumulatives forward from this day and update UI
                RecomputeCumulativesFrom(_activeDay.Date);
                UpdateAccountsDisplay();
                UpdateOverridePanels();
            }
            catch { }
        }

        // Compute whether a log entry is a manual reset anchor
        private static bool IsManualSet(AccountLogEntry a)
        {
            try
            {
                var n = a?.Note ?? string.Empty;
                return n.IndexOf("manual set", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch { return false; }
        }

        private IReadOnlyList<AccountLogEntry> GetLogs(DateTime from, DateTime to, string kind)
        {
            try { return TimeTrackingService.Instance.GetAccountLogEntries(from, to, kind); } catch { return Array.Empty<AccountLogEntry>(); }
        }

        // Compute starting running balance for a kind at a given start date using last manual set before start (if any), ignoring any auto-generated entries
        private double ComputeStartRunning(string kind, DateTime startDate)
        {
            double running = 0.0;
            try
            {
                var allBefore = TimeTrackingService.Instance.GetAccountLogEntries(DateTime.MinValue, startDate.AddDays(-1), kind)
                                                           .Where(a => (a.Note ?? string.Empty).IndexOf("manual", StringComparison.OrdinalIgnoreCase) >= 0)
                                                           .OrderBy(a => a.Date)
                                                           .ToList();
                var lastSet = allBefore.Where(IsManualSet).OrderBy(a => a.Date).LastOrDefault();
                if (lastSet != null)
                {
                    running = lastSet.Balance;
                    foreach (var e in allBefore.Where(x => x.Date > lastSet.Date && !IsManualSet(x))) running += e.Delta;
                }
                else
                {
                    var acc = TimeTrackingService.Instance.GetAccountState();
                    running = string.Equals(kind, "TIL", StringComparison.OrdinalIgnoreCase) ? acc.TILOffset : acc.HolidayOffset;
                    foreach (var e in allBefore.Where(x => !IsManualSet(x))) running += e.Delta;
                }
            }
            catch { }
            return running;
        }

        // New: recompute cumulative account balances across ALL known days using simple rules and manual anchors
        private List<DayViewModel> RecalculateAllDaysAndBalances()
        {
            var days = BuildAllDaysAcrossData();
            if (days == null || days.Count == 0) return days;

            var ordered = days.OrderBy(d => d.Date).ToList();
            var from = ordered.First().Date.Date;
            var to = ordered.Last().Date.Date;

            // Only consider manual log entries (manual set/delta) for anchors/adjustments
            var tilManual = GetLogs(from, to, "TIL").Where(a => (a.Note ?? string.Empty).IndexOf("manual", StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(a => a.Date).ToList();
            var holManual = GetLogs(from, to, "Holiday").Where(a => (a.Note ?? string.Empty).IndexOf("manual", StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(a => a.Date).ToList();

            double runningTIL = ComputeStartRunning("TIL", from);
            double runningHoliday = ComputeStartRunning("Holiday", from);

            foreach (var d in ordered)
            {
                // Apply same-day manual set anchors first
                var dayTil = tilManual.Where(a => a.Date.Date == d.Date.Date).OrderBy(a => a.Date).ToList();
                var dayHol = holManual.Where(a => a.Date.Date == d.Date.Date).OrderBy(a => a.Date).ToList();
                var tilSet = dayTil.LastOrDefault(IsManualSet);
                if (tilSet != null) runningTIL = tilSet.Balance;
                var holSet = dayHol.LastOrDefault(IsManualSet);
                if (holSet != null) runningHoliday = holSet.Balance;

                // Daily contributions (signed difference vs. target)
                double worked = d.Shifts.Sum(s => s.Hours);
                double target;
                if (d.TargetHours.HasValue) target = d.TargetHours.Value; else if (d.DayType == DayType.Vacation || d.DayType == DayType.PublicHoliday) target = 0; else if (d.Template != null && d.Template.HoursPerWeekday?.Length == 7) { int idx = ((int)d.Date.DayOfWeek + 6) % 7; target = d.Template.HoursPerWeekday[idx]; } else target = 0;
                var tilDelta = worked - target; // allow negative to decrease TIL
                var holDelta = (d.DayType == DayType.Vacation) ? -1.0 : 0.0;

                runningTIL += tilDelta;
                runningHoliday += holDelta;

                // Apply same-day manual delta adjustments after daily contributions
                foreach (var ml in dayTil.Where(x => !IsManualSet(x))) runningTIL += ml.Delta;
                foreach (var ml in dayHol.Where(x => !IsManualSet(x))) runningHoliday += ml.Delta;

                d.CumulativeTIL = runningTIL;
                d.CumulativeHoliday = runningHoliday;
            }

            return ordered;
        }

        // Recompute cumulatives for current month by recalculating across all days and mapping back
        private void RecomputeCumulatives()
        {
            try
            {
                var all = RecalculateAllDaysAndBalances();
                if (all == null || all.Count == 0 || _currentDays == null || _currentDays.Count == 0) return;
                var map = all.ToDictionary(d => d.Date.Date);
                var firstKnownDate = all.First().Date.Date;
                var lastKnownDate = all.Last().Date.Date;
                var lastKnownTIL = all.Last().CumulativeTIL;
                var lastKnownHol = all.Last().CumulativeHoliday;

                foreach (var d in _currentDays.OrderBy(x => x.Date))
                {
                    var key = d.Date.Date;
                    if (map.TryGetValue(key, out var src))
                    {
                        d.CumulativeTIL = src.CumulativeTIL;
                        d.CumulativeHoliday = src.CumulativeHoliday;
                    }
                    else if (key > lastKnownDate)
                    {
                        // Future date beyond known range: carry forward the last known balances
                        d.CumulativeTIL = lastKnownTIL;
                        d.CumulativeHoliday = lastKnownHol;
                    }
                    else if (key < firstKnownDate)
                    {
                        // Before known range: compute start running balances as of this date
                        d.CumulativeTIL = ComputeStartRunning("TIL", key.AddDays(1));
                        d.CumulativeHoliday = ComputeStartRunning("Holiday", key.AddDays(1));
                    }
                }
            }
            catch { }
        }

        // Recompute cumulatives forward from a specific date by recalculating all and mapping back
        private void RecomputeCumulativesFrom(DateTime startDate)
        {
            try
            {
                var all = RecalculateAllDaysAndBalances();
                if (all == null || all.Count == 0 || _currentDays == null || _currentDays.Count == 0) return;
                var map = all.ToDictionary(d => d.Date.Date);
                var firstKnownDate = all.First().Date.Date;
                var lastKnownDate = all.Last().Date.Date;
                var lastKnownTIL = all.Last().CumulativeTIL;
                var lastKnownHol = all.Last().CumulativeHoliday;

                foreach (var d in _currentDays.Where(x => x.Date.Date >= startDate.Date).OrderBy(x => x.Date))
                {
                    var key = d.Date.Date;
                    if (map.TryGetValue(key, out var src))
                    {
                        d.CumulativeTIL = src.CumulativeTIL;
                        d.CumulativeHoliday = src.CumulativeHoliday;
                    }
                    else if (key > lastKnownDate)
                    {
                        d.CumulativeTIL = lastKnownTIL;
                        d.CumulativeHoliday = lastKnownHol;
                    }
                    else if (key < firstKnownDate)
                    {
                        d.CumulativeTIL = ComputeStartRunning("TIL", key.AddDays(1));
                        d.CumulativeHoliday = ComputeStartRunning("Holiday", key.AddDays(1));
                    }
                }
                DaysList.Items.Refresh();
                MonthGrid.Items.Refresh();
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
                // Immediately set displayed total so it is visible even if UpdateShiftTotals silently fails
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
                        RecomputeCumulativesFrom(_activeDay.Date);
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
                    RecomputeCumulativesFrom(_activeDay.Date);
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
                        RecomputeCumulativesFrom(_activeDay.Date);
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
                        RecomputeCumulativesFrom(_activeDay.Date);
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
                var txt = new TextBox { Width = 160, Text = "0" };
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
                    var targetDate = _activeDay?.Date.Date ?? DateTime.Today.Date;

                    if (applyAsDelta)
                    {
                        var entry = new AccountLogEntry { Date = targetDate.AddHours(12), Kind = "Holiday", Delta = val, Balance = double.NaN, Note = "Manual delta", AffectedDate = targetDate };
                        TimeTrackingService.Instance.AddAccountLogEntry(entry);
                    }
                    else
                    {
                        var entry = new AccountLogEntry { Date = targetDate.AddHours(12), Kind = "Holiday", Delta = 0, Balance = val, Note = "Manual set", AffectedDate = targetDate };
                        TimeTrackingService.Instance.AddAccountLogEntry(entry);
                    }

                    // Reload persisted state to ensure any other components reading from disk see latest data
                    try { TimeTrackingService.Instance.Reload(); } catch { }

                    // Refresh computed cumulatives and UI so changes are visible immediately
                    try
                    {
                        PopulateDaysForCurrentMonth();
                        RecomputeCumulativesFrom(targetDate);
                        UpdateAccountsDisplay();
                        UpdateShiftTotalsDisplay();
                        UpdateOverridePanels();
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
                var txt = new TextBox { Width = 160, Text = "0" };
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
                    var targetDate = _activeDay?.Date.Date ?? DateTime.Today.Date;

                    if (applyAsDelta)
                    {
                        var entry = new AccountLogEntry { Date = targetDate.AddHours(12), Kind = "TIL", Delta = val, Balance = double.NaN, Note = "Manual delta", AffectedDate = targetDate };
                        TimeTrackingService.Instance.AddAccountLogEntry(entry);
                    }
                    else
                    {
                        var entry = new AccountLogEntry { Date = targetDate.AddHours(12), Kind = "TIL", Delta = 0, Balance = val, Note = "Manual set", AffectedDate = targetDate };
                        TimeTrackingService.Instance.AddAccountLogEntry(entry);
                    }

                    // Reload persisted state
                    try { TimeTrackingService.Instance.Reload(); } catch { }

                    // Refresh computed cumulatives and UI so changes are visible immediately
                    try
                    {
                        PopulateDaysForCurrentMonth();
                        RecomputeCumulativesFrom(targetDate);
                        UpdateAccountsDisplay();
                        UpdateShiftTotalsDisplay();
                        UpdateOverridePanels();
                    }
                    catch { }

                    MessageBox.Show("Time in Lieu account updated.", "Reset", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch { }
        }

        // Export year colour map (PDF)
        private void ExportYearMap_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var allDays = BuildAllDaysAcrossData();
                if (allDays == null || allDays.Count == 0)
                {
                    MessageBox.Show("No data available to export.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Gather available years and open combined options dialog
                var years = allDays.Select(d => d.Date.Year).Distinct().OrderBy(y => y).ToList();
                var dlg = new ExportOptionsDialog(years, new[] { DateTime.Today.Year }, defaultDayType: false, defaultLocation: true, defaultOvertime: true) { Owner = this };
                if (dlg.ShowDialog() != true) return;
                var selectedYears = dlg.SelectedYears;
                if (selectedYears == null || selectedYears.Count == 0)
                {
                    MessageBox.Show("Select at least one year.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                bool includeDayType = dlg.IncludeDayType;
                bool includeLocation = dlg.IncludeLocation;
                bool includeOvertime = dlg.IncludeOvertime;
                if (!includeDayType && !includeLocation && !includeOvertime)
                {
                    MessageBox.Show("Select at least one map type.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Output file
                var sfd = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "PDF files (*.pdf)|*.pdf",
                    FileName = "YearMap.pdf"
                };
                if (sfd.ShowDialog(this) != true) return;

                using (var doc = new PdfDocument())
                {
                    // Group pages by map type first (type -> all years), then move to next type
                    if (includeDayType)
                    {
                        foreach (var y in selectedYears)
                        {
                            var page = doc.AddPage();
                            using (var gfx = XGraphics.FromPdfPage(page))
                            {
                                DrawYearGridToGraphics(gfx, allDays, y, ExportMapMode.DayType);
                            }
                        }
                    }
                    if (includeLocation)
                    {
                        foreach (var y in selectedYears)
                        {
                            var page = doc.AddPage();
                            using (var gfx = XGraphics.FromPdfPage(page))
                            {
                                DrawYearGridToGraphics(gfx, allDays, y, ExportMapMode.PhysicalLocation);
                            }
                        }
                    }
                    if (includeOvertime)
                    {
                        foreach (var y in selectedYears)
                        {
                            var page = doc.AddPage();
                            using (var gfx = XGraphics.FromPdfPage(page))
                            {
                                DrawYearGridToGraphics(gfx, allDays, y, ExportMapMode.OvertimeGradient);
                            }
                        }
                    }

                    doc.Save(sfd.FileName);
                }

                MessageBox.Show("Exported year map to PDF.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to export PDF: " + ex.Message, "Export", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ColorModeCheck_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                ColorByPhysicalLocation = true;
                MonthGrid.Items.Refresh();
                DaysList.Items.Refresh();
            }
            catch { }
        }

        private void ColorModeCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                ColorByPhysicalLocation = false;
                MonthGrid.Items.Refresh();
                DaysList.Items.Refresh();
            }
            catch { }
        }

        private void DrawLegend(XGraphics g, double x, double y, ExportMapMode mode, List<DayViewModel> days, double availableWidth, int year)
        {
            var font = new XFont("Arial", 9, XFontStyle.Regular); double box = 12; double gap = 8; double cx = x; double startX = x; double padding = 6;
            if (mode == ExportMapMode.DayType)
            {
                var items = DayTypePdfColors.Keys.Select(k => k.ToString()).ToList();
                var widths = items.Select(it => g.MeasureString(it, font).Width + box + gap + padding).ToList(); double totalW = widths.Sum(); cx = startX + Math.Max(0, (availableWidth - totalW) / 2);
                for (int i = 0; i < items.Count; i++)
                {
                    var key = items[i]; var brush = new XSolidBrush(DayTypePdfColors[(DayType)Enum.Parse(typeof(DayType), key)]);
                    g.DrawRectangle(brush, cx, y, box, box); g.DrawString(key, font, XBrushes.Black, new XPoint(cx + box + 6, y + box - 2)); cx += widths[i];
                }
            }
            else if (mode == ExportMapMode.PhysicalLocation)
            {
                // Use persisted or generated-persisted colors and ensure wrapping/fit
                var locFont = new XFont("Arial", 8, XFontStyle.Regular);
                List<string> locs = days.Where(d => d.Date.Year == year)
                                         .Where(d => !string.IsNullOrWhiteSpace(d.PhysicalLocationOverride) || !string.IsNullOrWhiteSpace(d.LocationOverride) || !string.IsNullOrWhiteSpace(d.Template?.Location))
                                         .Select(d => !string.IsNullOrWhiteSpace(d.PhysicalLocationOverride) ? d.PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(d.LocationOverride) ? d.LocationOverride : d.Template?.Location ?? ""))
                                         .Where(s => !string.IsNullOrWhiteSpace(s))
                                         .GroupBy(s => s).OrderByDescending(gp => gp.Count()).Select(gp => gp.Key).ToList();
                int max = Math.Min(8, locs.Count);

                double curX = startX;
                double curY = y;
                double lineH = box + 6;

                for (int i = 0; i < max; i++)
                {
                    var key = locs[i];
                    var hex = TimeTrackingService.Instance.GetOrCreateLocationColor(key);
                    XBrush brush;
                    if (!string.IsNullOrWhiteSpace(hex))
                    {
                        try { var sysc = (Color)ColorConverter.ConvertFromString(hex); brush = new XSolidBrush(XColor.FromArgb(sysc.A, sysc.R, sysc.G, sysc.B)); }
                        catch { brush = XBrushes.White; }
                    }
                    else brush = XBrushes.White;

                    // Measure and shrink-to-fit if a single item is wider than available width
                    XFont itemFont = locFont;
                    double labelWidth = g.MeasureString(key, itemFont).Width;
                    double itemWidth = box + gap + labelWidth + padding;
                    if (itemWidth > availableWidth)
                    {
                        for (int sz = 7; sz >= 6; sz--)
                        {
                            var tryFont = new XFont("Arial", sz, XFontStyle.Regular);
                            var w = g.MeasureString(key, tryFont).Width + box + gap + padding;
                            if (w <= availableWidth)
                            {
                                itemFont = tryFont; itemWidth = w; break;
                            }
                        }
                    }

                    if (curX + itemWidth > startX + availableWidth)
                    {
                        curX = startX;
                        curY += lineH;
                    }

                    g.DrawRectangle(brush, curX, curY, box, box);
                    g.DrawString(key, itemFont, XBrushes.Black, new XPoint(curX + box + 6, curY + box - 2));
                    curX += itemWidth;
                }
            }
            else
            {
                // Overtime legend: red (-4h) .. white (0) .. green (+4h)
                double gradW = Math.Min(300, availableWidth - 160);
                double gradH = box;
                double startXGrad = x + (availableWidth - gradW) / 2;
                int steps = 60;
                for (int i = 0; i < steps; i++)
                {
                    double t = (double)i / (steps - 1); // 0..1
                    double delta = (t - 0.5) * 8.0; // -4..+4
                    var brush = GetOvertimeBrush(delta);
                    g.DrawRectangle(brush, startXGrad + i * (gradW / steps), y, gradW / steps + 0.5, gradH);
                }
                var small = new XFont("Arial", 8, XFontStyle.Regular);
                g.DrawString("-4h", small, XBrushes.Black, new XPoint(startXGrad - 20, y + box + 12));
                g.DrawString("0", small, XBrushes.Black, new XPoint(startXGrad + gradW / 2 - 4, y + box + 12));
                g.DrawString("+4h", small, XBrushes.Black, new XPoint(startXGrad + gradW + 4, y + box + 12));
            }
        }

        private XBrush GetOvertimeBrush(double delta)
        {
            double maxAbs = 4.0;
            double t;
            if (delta >= 0)
            {
                t = Math.Min(delta / maxAbs, 1.0);
                byte r = (byte)(255 - (int)(t * (255 - 0x2E)));
                byte g = (byte)(255 - (int)(t * (255 - 0xCC)));
                byte b = (byte)(255 - (int)(t * (255 - 0x71)));
                return new XSolidBrush(XColor.FromArgb(255, r, g, b));
            }
            else
            {
                t = Math.Min((-delta) / maxAbs, 1.0);
                byte r = (byte)(255 - (int)(t * (255 - 0xFF)));
                byte g = (byte)(255 - (int)(t * (255 - 0x6B)));
                byte b = (byte)(255 - (int)(t * (255 - 0x6B)));
                return new XSolidBrush(XColor.FromArgb(255, r, g, b));
            }
        }

        private void UpdateAccountsDisplay()
        {
            try
            {
                DayViewModel src = _activeDay;
                var today = DateTime.Today;
                if (src == null)
                    src = _currentDays?.FirstOrDefault(d => d.Date.Date == today.Date);

                if (src != null)
                {
                    if (this.FindName("AccountTILValue") is TextBlock til)
                        til.Text = FormatHoursAsHHmm(src.CumulativeTIL);
                    if (this.FindName("AccountHolidayValue") is TextBlock hol)
                        hol.Text = src.CumulativeHoliday.ToString("0.##") + " days";
                    return;
                }

                double tilBal = ComputeStartRunning("TIL", today.Date + TimeSpan.FromDays(1));
                double holBal = ComputeStartRunning("Holiday", today.Date + TimeSpan.FromDays(1));
                if (this.FindName("AccountTILValue") is TextBlock til2)
                    til2.Text = FormatHoursAsHHmm(tilBal);
                if (this.FindName("AccountHolidayValue") is TextBlock hol2)
                    hol2.Text = holBal.ToString("0.##") + " days";
            }
            catch { }
        }

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

                double delta = totalWorked - target;

                if (this.FindName("TotalWorkedDataTextBlock") is TextBlock tw2)
                    tw2.Text = FormatHoursAsHHmm(totalWorked);
                if (this.FindName("TotalDeltaDataTextBlock") is TextBlock td2)
                    td2.Text = FormatHoursAsHHmm(delta);
                if (this.FindName("ShiftTargetText") is TextBox st2)
                {
                    if (explicitOverride)
                        st2.Text = target.ToString("0.##");
                    else
                        st2.Text = target > 0 ? target.ToString("0.##") : "0";
                }
            }
            catch { }
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
                        var entry = new AccountLogEntry { Date = _activeDay.Date.AddHours(12), Kind = "TIL", Delta = diff, Balance = double.NaN, Note = "Manual override edit", AffectedDate = _activeDay.Date };
                        TimeTrackingService.Instance.AddAccountLogEntry(entry);
                        try { TimeTrackingService.Instance.Reload(); } catch { }
                        RecomputeCumulativesFrom(_activeDay.Date);
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
                var date = _activeDay.Date.Date;
                if (MessageBox.Show("Delete TIL manual overrides for this day?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;

                var removed = TimeTrackingService.Instance.RemoveManualAccountLogEntries(date, "TIL");
                try { TimeTrackingService.Instance.Reload(); } catch { }
                RecomputeCumulativesFrom(date);
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
                        var entry = new AccountLogEntry { Date = _activeDay.Date.AddHours(12), Kind = "Holiday", Delta = diff, Balance = double.NaN, Note = "Manual override edit", AffectedDate = _activeDay.Date };
                        TimeTrackingService.Instance.AddAccountLogEntry(entry);
                        try { TimeTrackingService.Instance.Reload(); } catch { }
                        RecomputeCumulativesFrom(_activeDay.Date);
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
                var date = _activeDay.Date.Date;
                if (MessageBox.Show("Delete Holiday manual overrides for this day?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;

                var removed = TimeTrackingService.Instance.RemoveManualAccountLogEntries(date, "Holiday");
                try { TimeTrackingService.Instance.Reload(); } catch { }
                RecomputeCumulativesFrom(date);
                UpdateAccountsDisplay();
                UpdateShiftTotalsDisplay();
                UpdateOverridePanels();
                DaysList.Items.Refresh();
                MonthGrid.Items.Refresh();
            }
            catch { }
        }

        // Restored helpers required by XAML and other methods
        public static string FormatHoursAsHHmm(double hours)
        {
            try
            {
                var ts = TimeSpan.FromHours(hours);
                var sign = ts.Ticks < 0 ? "-" : "";
                ts = new TimeSpan(Math.Abs(ts.Ticks));
                return sign + ((int)ts.TotalHours).ToString("00") + ":" + ts.Minutes.ToString("00");
            }
            catch { return "00:00"; }
        }

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

                // Holiday panel
                var holSet = GetManualSetForDay(date, "Holiday");
                double holDelta = GetManualOverrideSumForDay(date, "Holiday");
                if (holPanel != null && holLabel != null)
                {
                    if (holSet != null)
                    {
                        holPanel.Visibility = Visibility.Visible;
                        holLabel.Text = $"Manual set to: {holSet.Balance:0.##} d";
                    }
                    else if (Math.Abs(holDelta) > 0.0001)
                    {
                        holPanel.Visibility = Visibility.Visible;
                        holLabel.Text = $"Manual override: {holDelta:+0.##;-0.##;0} d";
                    }
                    else
                    {
                        holPanel.Visibility = Visibility.Collapsed;
                    }
                }

                // TIL panel
                var tilSet = GetManualSetForDay(date, "TIL");
                double tilDelta = GetManualOverrideSumForDay(date, "TIL");
                if (tilPanel != null && tilLabel != null)
                {
                    if (tilSet != null)
                    {
                        tilPanel.Visibility = Visibility.Visible;
                        tilLabel.Text = $"Manual set to: {FormatHoursAsHHmm(tilSet.Balance)}";
                    }
                    else if (Math.Abs(tilDelta) > 0.0001)
                    {
                        tilPanel.Visibility = Visibility.Visible;
                        tilLabel.Text = $"Manual override: {FormatHoursAsHHmm(tilDelta)}";
                    }
                    else
                    {
                        tilPanel.Visibility = Visibility.Collapsed;
                    }
                }
            }
            catch { }
        }

        private static AccountLogEntry GetManualSetForDay(DateTime date, string kind)
        {
            try
            {
                var items = TimeTrackingService.Instance.GetAccountLogEntries(date, date, kind);
                if (items == null) return null;
                return items.Where(a => (a.Note ?? string.Empty).IndexOf("manual set", StringComparison.OrdinalIgnoreCase) >= 0)
                            .OrderBy(a => a.Date)
                            .LastOrDefault();
            }
            catch { return null; }
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

        private void ImportExcel_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Import from Excel is not implemented yet.", "Import", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Persist any pending edits so recalculation sees latest
                SaveCurrentDayEdits();

                // Recalculate across all known days using the same logic as the UI
                var all = RecalculateAllDaysAndBalances();
                if (all == null || all.Count == 0)
                {
                    MessageBox.Show("No data to export.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                var ordered = all.OrderBy(d => d.Date).ToList();

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

                    // cache manual reset log entries by date for the full exported range
                    var from = ordered.First().Date.Date;
                    var to = ordered.Last().Date.Date;
                    var allLog = TimeTrackingService.Instance.GetAccountLogEntries(from, to, null);

                    int row = 2;
                    foreach (var d in ordered)
                    {
                        ws.Cell(row, 1).Value = d.Date; ws.Cell(row, 1).Style.DateFormat.Format = "yyyy-MM-dd";
                        ws.Cell(row, 2).Value = d.Date.ToString("dddd");
                        ws.Cell(row, 3).Value = d.PositionOverride ?? d.Template?.JobDescription ?? string.Empty;
                        ws.Cell(row, 4).Value = d.LocationOverride ?? d.Template?.Location ?? string.Empty;
                        ws.Cell(row, 5).Value = d.PhysicalLocationOverride ?? (!string.IsNullOrWhiteSpace(d.LocationOverride) ? d.LocationOverride : d.Template?.Location ?? string.Empty);
                        ws.Cell(row, 6).Value = d.DayType.ToString();
                        ws.Cell(row, 7).Value = d.TargetComputed; ws.Cell(row, 7).Style.NumberFormat.Format = "0.00";
                        ws.Cell(row, 8).Value = d.Worked; ws.Cell(row, 8).Style.NumberFormat.Format = "0.00";
                        double tilDelta = d.Worked - d.TargetComputed;
                        ws.Cell(row, 9).Value = tilDelta; ws.Cell(row, 9).Style.NumberFormat.Format = "0.00";
                        double deltaHoliday = d.DayType == DayType.Vacation ? -1.0 : 0.0;
                        ws.Cell(row, 10).Value = deltaHoliday; ws.Cell(row, 10).Style.NumberFormat.Format = "0.00";
                        ws.Cell(row, 11).Value = d.CumulativeTIL; ws.Cell(row, 11).Style.NumberFormat.Format = "0.00";
                        ws.Cell(row, 12).Value = d.CumulativeHoliday; ws.Cell(row, 12).Style.NumberFormat.Format = "0.00";

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
                            int c = 16 + i * 4;
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

                    // Manual Resets sheet for full range
                    var logWs = wb.Worksheets.Add("ManualResets");
                    logWs.Cell(1, 1).Value = "Date"; logWs.Cell(1, 2).Value = "Kind"; logWs.Cell(1, 3).Value = "Delta"; logWs.Cell(1, 4).Value = "Balance"; logWs.Cell(1, 5).Value = "Note"; logWs.Cell(1, 6).Value = "AffectedDate"; logWs.Row(1).Style.Font.Bold = true;
                    int lr = 2; foreach (var a in TimeTrackingService.Instance.GetAccountLogEntries(from, to, null).Where(a => (a.Note ?? string.Empty).IndexOf("manual", StringComparison.OrdinalIgnoreCase) >= 0))
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

        private List<DayViewModel> BuildAllDaysAcrossData()
        {
            var days = new List<DayViewModel>();
            try
            {
                DateTime? min = null; DateTime? max = null;
                try { var events = TimeTrackingService.Instance.GetEvents(); if (events != null && events.Count > 0) { var mn = events.Min(e => e.Timestamp.Date); var mx = events.Max(e => e.Timestamp.Date); min = min.HasValue ? (mn < min ? mn : min) : mn; max = max.HasValue ? (mx > max ? mx : mx) : mx; } } catch { }
                try { var sh = TimeTrackingService.Instance.GetAllSavedShifts(); if (sh != null && sh.Count > 0) { var mn = sh.Min(s => s.Date.Date); var mx = sh.Max(s => s.Date); min = min.HasValue ? (mn < min ? mn : min) : mn; max = max.HasValue ? (mx > max ? mx : mx) : mx; } } catch { }
                try { var ov = TimeTrackingService.Instance.GetOverrides(); if (ov != null && ov.Count > 0) { var mn = ov.Min(o => o.Date.Date); var mx = ov.Max(o => o.Date.Date); min = min.HasValue ? (mn < min ? mn : min) : mn; max = max.HasValue ? (mx > max ? mx : mx) : mx; } } catch { }
                try { var tmps = TimeTrackingService.Instance.GetTemplates(); if (tmps != null && tmps.Count > 0) { var mn = tmps.Min(t => t.StartDate.Date); var mx = tmps.Max(t => (t.EndDate ?? DateTime.Today).Date); min = min.HasValue ? (mn < min ? mn : min) : mn; max = max.HasValue ? (mx > max ? mx : mx) : mx; } } catch { }
                if (!min.HasValue || !max.HasValue)
                {
                    var f = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1); var l = f.AddMonths(1).AddDays(-1); var summaries0 = TimeTrackingService.Instance.GetDaySummaries(f, l); days = summaries0.Select(s => new DayViewModel(s)).ToList();
                }
                else { var summaries = TimeTrackingService.Instance.GetDaySummaries(min.Value, max.Value); days = summaries.Select(s => new DayViewModel(s)).ToList(); }

                var ordered = days.OrderBy(d => d.Date).ToList();
                foreach (var d in ordered)
                {
                    var ov = TimeTrackingService.Instance.GetOverrideForDate(d.Date);
                    if (ov != null)
                    {
                        d.PositionOverride = ov.Position; d.LocationOverride = ov.Location; d.PhysicalLocationOverride = string.IsNullOrWhiteSpace(ov.PhysicalLocation) ? (string.IsNullOrWhiteSpace(ov.Location) ? d.Template?.Location : ov.Location) : ov.PhysicalLocation; d.TargetHours = ov.TargetHours; if (ov.DayType.HasValue) d.SetDayType(ov.DayType.Value);
                    }
                    else if (d.Template != null && !string.IsNullOrWhiteSpace(d.Template.Location)) d.PhysicalLocationOverride = d.Template.Location;
                    var shs = TimeTrackingService.Instance.GetShiftsForDate(d.Date).ToList(); d.Shifts = shs.Any() ? shs.Select(s2 => new Shift { Date = s2.Date, Start = s2.Start, End = s2.End, Description = s2.Description, DayMode = s2.DayMode, ManualStartOverride = s2.ManualStartOverride, ManualEndOverride = s2.ManualEndOverride, LunchBreak = s2.LunchBreak }).ToList() : new List<Shift>();
                }
                days = ordered;
            }
            catch { }
            return days;
        }

        private void DrawYearGridToGraphics(XGraphics g, List<DayViewModel> days, int year, ExportMapMode mode)
        {
            double margin = 30; double pageWidth = g.PageSize.Width; double pageHeight = g.PageSize.Height; double contentWidth = pageWidth - (margin * 2); double legendHeight = 80;
            var headerFont = new XFont("Arial", 18, XFontStyle.Bold);
            var sub = mode == ExportMapMode.PhysicalLocation ? "Physical location map - " : mode == ExportMapMode.OvertimeGradient ? "Overtime map - " : "Day type map - ";
            g.DrawString(sub + year.ToString(), headerFont, XBrushes.Black, new XRect(margin, 10, contentWidth, 30), XStringFormats.TopLeft);
            int cols = 3; int rows = 4; double gap = 8; double monthWidth = (contentWidth - (gap * (cols - 1))) / cols; double monthGridAvailableHeight = pageHeight - 80 - legendHeight - (gap * (rows - 1)); double monthHeight = monthGridAvailableHeight / rows;
            for (int m = 1; m <= 12; m++)
            {
                int monthIndex = m - 1; int col = monthIndex % cols; int row = monthIndex / cols; double x = margin + col * (monthWidth + gap); double y = 50 + row * (monthHeight + gap);
                DrawSingleMonth(g, days, year, m, new XRect(x, y, monthWidth, monthHeight), mode);
            }
            DrawLegend(g, margin, pageHeight - legendHeight + 12, mode, days, contentWidth, year);
        }
        private void DrawSingleMonth(XGraphics g, List<DayViewModel> days, int year, int month, XRect rect, ExportMapMode mode)
        {
            var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(month);
            var titleFont = new XFont("Arial", 12, XFontStyle.Bold);
            g.DrawString(monthName + " " + year.ToString(), titleFont, XBrushes.Black, new XPoint(rect.X + 4, rect.Y + 14));
            double top = rect.Y + 22;
            double left = rect.X + 2;
            double gridWidth = rect.Width - 4;
            double gridHeight = rect.Height - 28;
            int cols = 7; int rows = 6;
            double cellW = gridWidth / cols; double cellH = gridHeight / rows;
            var first = new DateTime(year, month, 1);
            int leading = (int)first.DayOfWeek - 1; if (leading < 0) leading += 7;

            for (int i = 0; i < rows * cols; i++)
            {
                int dayNum = i - leading + 1;
                double cx = left + (i % cols) * cellW; double cy = top + (i / cols) * cellH;
                var cellRect = new XRect(cx, cy, cellW - 1, cellH - 1);
                if (dayNum >= 1 && dayNum <= DateTime.DaysInMonth(year, month))
                {
                    var d = days.FirstOrDefault(dd => dd.Date.Year == year && dd.Date.Month == month && dd.Date.Day == dayNum);
                    XBrush brush = XBrushes.White;
                    if (d != null)
                    {
                        if (mode == ExportMapMode.PhysicalLocation)
                        {
                            var key = !string.IsNullOrWhiteSpace(d.PhysicalLocationOverride) ? d.PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(d.LocationOverride) ? d.LocationOverride : d.Template?.Location ?? string.Empty);
                            if (string.IsNullOrWhiteSpace(key)) brush = XBrushes.White;
                            else
                            {
                                var hex = TimeTrackingService.Instance.GetOrCreateLocationColor(key);
                                if (!string.IsNullOrWhiteSpace(hex))
                                {
                                    try { var sysc = (Color)ColorConverter.ConvertFromString(hex); brush = new XSolidBrush(XColor.FromArgb(sysc.A, sysc.R, sysc.G, sysc.B)); }
                                    catch { brush = XBrushes.White; }
                                }
                                else brush = XBrushes.White;
                            }
                        }
                        else if (mode == ExportMapMode.OvertimeGradient)
                        {
                            bool isSpecial = d.DayType == DayType.Weekend || d.DayType == DayType.PublicHoliday || d.DayType == DayType.Vacation;
                            if (isSpecial && Math.Abs(d.Worked) < 0.0001)
                            {
                                brush = new XSolidBrush(XColor.FromArgb(0xFF, 0xF0, 0xF0, 0xF0));
                            }
                            else
                            {
                                double delta = d.Worked - d.TargetComputed;
                                brush = GetOvertimeBrush(delta);
                            }
                        }
                        else
                        {
                            if (DayTypePdfColors.TryGetValue(d.DayType, out var c)) brush = new XSolidBrush(c); else brush = XBrushes.White;
                        }
                    }
                    g.DrawRectangle(brush, cellRect);
                    var numFont = new XFont("Arial", 8, XFontStyle.Regular);
                    g.DrawString(dayNum.ToString(), numFont, XBrushes.Black, new XPoint(cellRect.X + 4, cellRect.Y + 10));
                }
                else
                {
                    g.DrawRectangle(XBrushes.LightGray, cellRect);
                }
            }
            g.DrawRectangle(XPens.Black, rect.X, rect.Y, rect.Width, rect.Height);
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
    }

    // View model for a day used by the TimeTrackingDialog. Provides values and formatting used by XAML bindings.
    public class DayViewModel
    {
        public DateTime Date { get; set; }

        // Template that applies to this day, if any
        public TimeTemplate Template { get; set; }

        // Per-day overrides (persisted)
        public string PositionOverride { get; set; }
        public string LocationOverride { get; set; }
        public string PhysicalLocationOverride { get; set; }
        public double? TargetHours { get; set; }

        // Working state
        public DayType DayType { get; private set; }
        public List<Shift> Shifts { get; set; } = new List<Shift>();

        // Selection flags used to style month grid
        public bool IsToday { get; set; }
        public bool IsSelected { get; set; }

        // Accounts
        public double CumulativeTIL { get; set; }
        public double CumulativeHoliday { get; set; }

        public DayViewModel(DaySummary s)
        {
            Date = s.Date.Date;
            // Default day type: weekend vs working day
            DayType = (Date.DayOfWeek == DayOfWeek.Saturday || DayOfWeek.Sunday == Date.DayOfWeek) ? DayType.Weekend : DayType.WorkingDay;
            // Find template for this date
            try { Template = TimeTrackingService.Instance.GetTemplates().FirstOrDefault(t => t.AppliesTo(Date)); } catch { }
        }

        public void SetDayType(DayType dt)
        {
            DayType = dt;
        }

        public string DateString => Date.ToString("dd MMM yyyy");
        public string Weekday => Date.ToString("dddd");

        // Computed totals
        public double TargetComputed
        {
            get
            {
                if (TargetHours.HasValue) return TargetHours.Value;
                if (DayType == DayType.Vacation || DayType == DayType.PublicHoliday) return 0;
                if (Template != null && Template.HoursPerWeekday != null && Template.HoursPerWeekday.Length == 7)
                {
                    int idx = (((int)Date.DayOfWeek + 6) % 7); // Mon=0..Sun=6
                    return Template.HoursPerWeekday[idx];
                }
                return 0;
            }
        }

        public double Worked => (Shifts != null && Shifts.Count > 0) ? Shifts.Sum(s => s.Hours) : 0.0;
        public double Delta
        {
            get
            {
                var target = TargetComputed;
                // Signed delta so TIL decreases on shortfall
                return Worked - target;
            }
        }

        public string WorkedDisplay => FormatHoursAsHHmmLocal(Worked);
        public string TargetDisplay => FormatHoursAsHHmmLocal(TargetComputed);

        private static string FormatHoursAsHHmmLocal(double hours)
        {
            try
            {
                var ts = TimeSpan.FromHours(hours);
                var sign = ts.Ticks < 0 ? "-" : "";
                ts = new TimeSpan(Math.Abs(ts.Ticks));
                return sign + ((int)ts.TotalHours).ToString("00") + ":" + ts.Minutes.ToString("00");
            }
            catch { return "00:00"; }
        }

        // Background for month grid item
        public Brush BackgroundBrush
        {
            get
            {
                try
                {
                    // New rule: if weekend, public holiday or vacation AND 0 hours worked, show grey
                    bool isSpecialDay = DayType == DayType.Weekend || DayType == DayType.PublicHoliday || DayType == DayType.Vacation;
                    if (isSpecialDay && Math.Abs(Worked) < 0.0001)
                    {
                        return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0F0F0"));
                    }

                    if (TimeTrackingDialog.ColorByPhysicalLocation)
                    {
                        var key = !string.IsNullOrWhiteSpace(PhysicalLocationOverride) ? PhysicalLocationOverride : (!string.IsNullOrWhiteSpace(LocationOverride) ? LocationOverride : Template?.EmploymentLocation ?? string.Empty);
                        if (string.IsNullOrWhiteSpace(key)) return Brushes.White;
                        var hex = TimeTrackingService.Instance.GetLocationColor(key);
                        if (!string.IsNullOrWhiteSpace(hex))
                        {
                            var c = (Color)ColorConverter.ConvertFromString(hex);
                            return new SolidColorBrush(c);
                        }
                        // fallback: deterministic color by hash
                        int h = key.GetHashCode();
                        byte r = (byte)(80 + (Math.Abs(h) % 176));
                        byte gcol = (byte)(80 + (Math.Abs(h / 7) % 176));
                        byte b = (byte)(80 + (Math.Abs(h / 13) % 176));
                        return new SolidColorBrush(Color.FromArgb(255, r, gcol, b));
                    }
                    // Color by day type
                    switch (DayType)
                    {
                        case DayType.WorkingDay: return (Brush)new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DFF0D8"));
                        case DayType.Weekend: return (Brush)new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0F0F0"));
                        case DayType.PublicHoliday: return (Brush)new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFE5E5"));
                        case DayType.Vacation: return (Brush)new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DDEEFF"));
                        case DayType.TimeInLieu: return (Brush)new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF2CC"));
                        case DayType.Other: return (Brush)new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EFEFEF"));
                        default: return null;
                    }
                }
                catch { return Brushes.White; }
            }
        }
    }

    // Persisted window placement
    public class WindowSettings
    {
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public WindowState State { get; set; }
    }
}

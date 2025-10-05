using System;

namespace Todo
{
    public class Shift
    {
        public DateTime Date { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public string Description { get; set; } = string.Empty;
        // Subtract lunch break from total hours
        public double Hours => Math.Max(0, (End - Start - LunchBreak).TotalHours);

        // Display properties used in UI
        public string StartDisplay => Start.ToString(@"hh\:mm");
        public string EndDisplay => End.ToString(@"hh\:mm");
        public string Display => string.Format("{0} - {1} ({2:0.00}h) {3}", StartDisplay, EndDisplay, Hours, Description);

        // Additional helpers for hh:mm display
        public TimeSpan WorkedTime => (End > Start) ? (End - Start - LunchBreak) : TimeSpan.Zero;
        public string WorkedHhmm => (WorkedTime < TimeSpan.Zero ? TimeSpan.Zero : WorkedTime).ToString(@"hh\:mm");
        public string LunchHhmm => LunchBreak.ToString(@"hh\:mm");

        // Manual override flags
        public bool ManualStartOverride { get; set; }
        public bool ManualEndOverride { get; set; }

        // New: lunch break and day mode
        public TimeSpan LunchBreak { get; set; }
        public string DayMode { get; set; }
        public string LunchDisplay => LunchBreak != TimeSpan.Zero ? "Lunch: " + LunchBreak.ToString(@"hh\:mm") : string.Empty;
        public string DayModeDisplay => !string.IsNullOrWhiteSpace(DayMode) ? $"Mode: {DayMode}" : string.Empty;

        // Computed helper for UI binding
        public bool IsManual => ManualStartOverride || ManualEndOverride;
    }
}

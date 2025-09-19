using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace Todo
{
    public enum TimeEventType { Open, Close }

    public class TimeEvent
    {
        public DateTime Timestamp { get; set; }
        public TimeEventType Type { get; set; }
    }

    public class TimeTemplate
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public DateTime StartDate { get; set; } = DateTime.Today;
        public DateTime? EndDate { get; set; } = null; // null = ongoing
        public string JobDescription { get; set; } = "";
        public string Location { get; set; } = "";
        // Hours for Monday(0) .. Sunday(6)
        public double[] HoursPerWeekday { get; set; } = new double[7] {8,8,8,8,8,0,0};

        public bool AppliesTo(DateTime d)
        {
            var date = d.Date;
            if (date < StartDate.Date) return false;
            if (EndDate.HasValue && date > EndDate.Value.Date) return false;
            return true;
        }
    }

    public class DaySummary
    {
        public DateTime Date { get; set; }
        public TimeSpan? OpenTime { get; set; }
        public TimeSpan? CloseTime { get; set; }
        public double? WorkedHours { get; set; }
        public double StandardHours { get; set; }
        public double? DeltaHours { get; set; }
    }

    public class TimeTrackingService
    {
        private const string EventsFile = "timetracking_events.json";
        private const string TemplatesFile = "timetracking_templates.json";
        private List<TimeEvent> _events = new List<TimeEvent>();
        private List<TimeTemplate> _templates = new List<TimeTemplate>();

        public static TimeTrackingService Instance { get; } = new TimeTrackingService();

        private TimeTrackingService()
        {
            Load();
        }

        private void Load()
        {
            try
            {
                if (File.Exists(EventsFile))
                {
                    var j = File.ReadAllText(EventsFile);
                    var items = JsonSerializer.Deserialize<List<TimeEvent>>(j);
                    if (items != null) _events = items.OrderBy(x=>x.Timestamp).ToList();
                }
            }
            catch { }

            try
            {
                if (File.Exists(TemplatesFile))
                {
                    var j = File.ReadAllText(TemplatesFile);
                    var items = JsonSerializer.Deserialize<List<TimeTemplate>>(j);
                    if (items != null) _templates = items;
                }
            }
            catch { }
        }

        private void SaveEvents()
        {
            try
            {
                var j = JsonSerializer.Serialize(_events, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(EventsFile, j);
            }
            catch { }
        }

        private void SaveTemplates()
        {
            try
            {
                var j = JsonSerializer.Serialize(_templates, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(TemplatesFile, j);
            }
            catch { }
        }

        public void RecordOpen()
        {
            try
            {
                _events.Add(new TimeEvent { Timestamp = DateTime.Now, Type = TimeEventType.Open });
                SaveEvents();
            }
            catch { }
        }

        public void RecordClose()
        {
            try
            {
                _events.Add(new TimeEvent { Timestamp = DateTime.Now, Type = TimeEventType.Close });
                SaveEvents();
            }
            catch { }
        }

        public IReadOnlyList<TimeEvent> GetEvents() => _events.AsReadOnly();

        public IReadOnlyList<TimeTemplate> GetTemplates() => _templates.AsReadOnly();

        public void UpsertTemplate(TimeTemplate t)
        {
            var found = _templates.FirstOrDefault(x => x.Id == t.Id);
            if (found == null)
            {
                _templates.Add(t);
            }
            else
            {
                found.StartDate = t.StartDate;
                found.EndDate = t.EndDate;
                found.JobDescription = t.JobDescription;
                found.Location = t.Location;
                found.HoursPerWeekday = t.HoursPerWeekday;
            }
            SaveTemplates();
        }

        public List<DaySummary> GetDaySummaries(DateTime from, DateTime to)
        {
            var result = new List<DaySummary>();
            // events already ordered
            var events = _events.Where(e => e.Timestamp.Date >= from.Date && e.Timestamp.Date <= to.Date).OrderBy(e => e.Timestamp).ToList();

            // For each date in range, find first Open on that date and last Close on that date
            for (var d = from.Date; d <= to.Date; d = d.AddDays(1))
            {
                var opens = _events.Where(e => e.Type == TimeEventType.Open && e.Timestamp.Date == d).OrderBy(e => e.Timestamp).ToList();
                var closes = _events.Where(e => e.Type == TimeEventType.Close && e.Timestamp.Date == d).OrderBy(e => e.Timestamp).ToList();
                TimeSpan? openTs = null;
                TimeSpan? closeTs = null;
                double? worked = null;
                if (opens.Count > 0 && closes.Count > 0)
                {
                    openTs = opens.First().Timestamp.TimeOfDay;
                    closeTs = closes.Last().Timestamp.TimeOfDay;
                    var openDt = opens.First().Timestamp;
                    var closeDt = closes.Last().Timestamp;
                    if (closeDt > openDt)
                        worked = (closeDt - openDt).TotalHours;
                }

                // find applicable template and standard hours
                double std = 0;
                var temp = _templates.FirstOrDefault(t => t.AppliesTo(d));
                if (temp != null)
                {
                    // Map DayOfWeek to index (Mon=0..Sun=6)
                    int idx = ((int)d.DayOfWeek + 6) % 7; // Monday=0
                    if (temp.HoursPerWeekday != null && temp.HoursPerWeekday.Length == 7)
                        std = temp.HoursPerWeekday[idx];
                }

                result.Add(new DaySummary
                {
                    Date = d,
                    OpenTime = openTs,
                    CloseTime = closeTs,
                    WorkedHours = worked,
                    StandardHours = std,
                    DeltaHours = worked.HasValue ? (worked.Value - std) : (double?)null
                });
            }

            return result;
        }
    }
}

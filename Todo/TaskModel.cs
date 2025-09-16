using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Todo
{
    [Serializable]
    public class TaskModel : INotifyPropertyChanged
    {
        private Guid _id = Guid.NewGuid();
        public Guid Id
        {
            get => _id;
            set
            {
                if (_id == value) return;
                _id = value;
                OnPropertyChanged(nameof(Id));
            }
        }

        private string _taskName = string.Empty;
        public string TaskName
        {
            get => _taskName;
            set
            {
                if (_taskName == value) return;
                _taskName = value;
                OnPropertyChanged(nameof(TaskName));
            }
        }

        private bool _isComplete = false;
        public bool IsComplete
        {
            get => _isComplete;
            set
            {
                if (_isComplete == value) return;
                _isComplete = value;
                OnPropertyChanged(nameof(IsComplete));
            }
        }

        private bool _isPlaceholder = false;
        public bool IsPlaceholder
        {
            get => _isPlaceholder;
            set
            {
                if (_isPlaceholder == value) return;
                _isPlaceholder = value;
                OnPropertyChanged(nameof(IsPlaceholder));
            }
        }

        private bool _isReadOnly = false;
        public bool IsReadOnly
        {
            get => _isReadOnly;
            set
            {
                if (_isReadOnly == value) return;
                _isReadOnly = value;
                OnPropertyChanged(nameof(IsReadOnly));
            }
        }

        private string _description = string.Empty;
        public string Description
        {
            get => _description;
            set
            {
                if (_description == value) return;
                _description = value;
                OnPropertyChanged(nameof(Description));
            }
        }

        private List<string> _people = new List<string>();
        public List<string> People
        {
            get => _people;
            set
            {
                if (_people == value) return;
                _people = value;
                OnPropertyChanged(nameof(People));
            }
        }

        private List<string> _meetings = new List<string>();
        public List<string> Meetings
        {
            get => _meetings;
            set
            {
                if (_meetings == value) return;
                _meetings = value;
                OnPropertyChanged(nameof(Meetings));
            }
        }

        private bool _isFuture = false;
        public bool IsFuture
        {
            get => _isFuture;
            set
            {
                if (_isFuture == value) return;
                _isFuture = value;
                OnPropertyChanged(nameof(IsFuture));
            }
        }

        private DateTime? _futureDate = null;
        public DateTime? FutureDate
        {
            get => _futureDate;
            set
            {
                if (_futureDate == value) return;
                _futureDate = value;
                OnPropertyChanged(nameof(FutureDate));
            }
        }

        // New: the date/key this task is associated with (yyyy-MM-dd)
        private DateTime? _setDate = null;
        public DateTime? SetDate
        {
            get => _setDate;
            set
            {
                if (_setDate == value) return;
                _setDate = value;
                OnPropertyChanged(nameof(SetDate));
            }
        }

        // New: whether the UI should show the date for this task (used in All view)
        private bool _showDate = false;
        public bool ShowDate
        {
            get => _showDate;
            set
            {
                if (_showDate == value) return;
                _showDate = value;
                OnPropertyChanged(nameof(ShowDate));
            }
        }

        public override string ToString()
        {
            string completeAsString = "Complete";

            if (!IsComplete)
                completeAsString = "Incomplete";

            return $"TASK {TaskName} ({completeAsString})";
        }

        public TaskModel()
        {

        }

        public TaskModel(string taskname, bool iscomplete = false, bool isPlaceholder = false, string description = "", List<string> people = null, List<string> meetings = null, bool isFuture = false, DateTime? futureDate = null, Guid? id = null)
        {
            Id = id ?? Guid.NewGuid();
            TaskName = taskname;
            IsComplete = iscomplete;
            IsPlaceholder = isPlaceholder;
            Description = description ?? string.Empty;
            People = people ?? new List<string>();
            Meetings = meetings ?? new List<string>();
            IsFuture = isFuture;
            FutureDate = futureDate;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

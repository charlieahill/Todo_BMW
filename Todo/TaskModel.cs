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

        public TaskModel(string taskname, bool iscomplete = false, bool isPlaceholder = false)
        {
            TaskName = taskname;
            IsComplete = iscomplete;
            IsPlaceholder = isPlaceholder;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

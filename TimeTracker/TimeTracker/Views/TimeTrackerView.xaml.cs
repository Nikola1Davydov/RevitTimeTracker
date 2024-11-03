using TimeTracker.ViewModels;

namespace TimeTracker.Views
{
    public sealed partial class TimeTrackerView
    {
        public TimeTrackerView(TimeTrackerViewModel viewModel)
        {
            DataContext = viewModel;
            InitializeComponent();
        }
    }
}
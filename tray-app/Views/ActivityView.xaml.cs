using System.Windows;
using System.Windows.Controls;
using MediaPipelineTray.ViewModels;

namespace MediaPipelineTray.Views;

public partial class ActivityView : UserControl
{
    public ActivityView() => InitializeComponent();

    private MainViewModel? Model => DataContext as MainViewModel;

    private void OnPauseLane(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: LaneRow lane })
        {
            Model?.PauseLane(lane);
        }
    }

    private void OnOpenFailedFolder(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: FailureRow failure })
        {
            Model?.OpenLaneFolder(failure.Preset, failure.Workspace);
        }
    }

    private void OnRequeue(object sender, RoutedEventArgs e)
    {
        if (Model is null)
        {
            return;
        }

        var moved = Model.RequeueFailures();

        MessageBox.Show(
            Window.GetWindow(this),
            moved == 0
                ? "Nothing to requeue. The failed folder is empty."
                : $"Moved {moved} file(s) back to the input folder.",
            "Requeue",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void OnDismissFailures(object sender, RoutedEventArgs e) => Model?.DismissFailures();
}

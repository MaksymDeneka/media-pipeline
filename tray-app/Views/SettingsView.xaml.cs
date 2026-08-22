using System.Windows;
using System.Windows.Controls;
using MediaPipelineTray.ViewModels;

namespace MediaPipelineTray.Views;

public partial class SettingsView : UserControl
{
    public SettingsView() => InitializeComponent();

    private SettingsViewModel? Model => DataContext as SettingsViewModel;

    private void OnDiscard(object sender, RoutedEventArgs e) => Model?.Load();

    private async void OnSave(object sender, RoutedEventArgs e)
    {
        if (Model is null || sender is not FrameworkElement button)
        {
            return;
        }

        button.IsEnabled = false;

        try
        {
            var restarted = await Model.SaveAndRestartAsync();
            Model.Load();

            if (!restarted)
            {
                MessageBox.Show(
                    Window.GetWindow(this),
                    "Settings were saved, but the watcher did not restart within 90 seconds. It "
                    + "finishes the file it is on before exiting, so a long encode can take a "
                    + "while. Nothing was killed, and the new settings apply once it comes back.",
                    "Saved, restart pending",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }
        finally
        {
            button.IsEnabled = true;
        }
    }
}

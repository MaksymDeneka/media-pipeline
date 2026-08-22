using System.IO;
using System.Windows;
using System.Windows.Controls;
using MediaPipelineTray.ViewModels;

namespace MediaPipelineTray.Views;

public partial class PresetsView : UserControl
{
    public PresetsView() => InitializeComponent();

    private PresetsViewModel? Model => DataContext as PresetsViewModel;

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

    private void OnAddPreset(object sender, RoutedEventArgs e)
    {
        if (Model is null)
        {
            return;
        }

        var name = PromptDialog.Ask(
            Window.GetWindow(this),
            "New preset",
            "Name it after what it produces, for example photos or clips. This becomes its "
            + "folder name, so keep it short and avoid spaces.");

        if (name is null)
        {
            return;
        }

        name = name.Trim();

        if (name.Length == 0 || name.Any(c => Path.GetInvalidFileNameChars().Contains(c)))
        {
            MessageBox.Show(
                Window.GetWindow(this),
                "That name cannot be used as a folder name.",
                "Invalid name",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        try
        {
            Model.AddPreset(name);
        }
        catch (InvalidOperationException ex)
        {
            MessageBox.Show(Window.GetWindow(this), ex.Message, "Cannot add preset",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void OnRemovePreset(object sender, RoutedEventArgs e)
    {
        if (Model?.Selected is not { } preset)
        {
            return;
        }

        var answer = MessageBox.Show(
            Window.GetWindow(this),
            $"Remove the '{preset.Name}' preset from the configuration?\n\n"
            + "Its folders and any media in them are left exactly as they are. Only the "
            + "settings entry is removed.",
            "Remove preset",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Question);

        if (answer == MessageBoxResult.OK)
        {
            Model.RemovePreset(preset.Name);
        }
    }
}

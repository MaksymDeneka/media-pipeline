using System.Windows;

namespace MediaPipelineTray.Views;

/// <summary>A single-line text prompt, since WPF has no built-in equivalent of an input box.</summary>
public partial class PromptDialog : Window
{
    private PromptDialog() => InitializeComponent();

    /// <summary>Returns the entered text, or null when cancelled.</summary>
    public static string? Ask(Window? owner, string heading, string explanation, string initial = "")
    {
        var dialog = new PromptDialog { Owner = owner };
        dialog.Heading.Text = heading;
        dialog.Explanation.Text = explanation;
        dialog.Answer.Text = initial;
        dialog.Answer.Focus();
        dialog.Answer.SelectAll();

        return dialog.ShowDialog() == true ? dialog.Answer.Text : null;
    }

    private void OnAccept(object sender, RoutedEventArgs e) => DialogResult = true;

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}

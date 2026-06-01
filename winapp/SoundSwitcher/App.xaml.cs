using System.Windows;
using System.Windows.Threading;

namespace SoundSwitcher;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += (s, ex) =>
        {
            System.IO.File.AppendAllText(
                System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SoundSwitcher_crash.txt"),
                $"{DateTime.Now}: {ex.Exception}\n\n");
            MessageBox.Show(ex.Exception.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            ex.Handled = true;
        };
        AppDomain.CurrentDomain.UnhandledException += (s, ex) =>
        {
            System.IO.File.AppendAllText(
                System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SoundSwitcher_crash.txt"),
                $"{DateTime.Now}: {ex.ExceptionObject}\n\n");
        };
    }
}

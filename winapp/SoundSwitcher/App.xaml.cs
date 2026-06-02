using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace SoundSwitcher;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        // Force software rendering. This static, non-animated utility window does not
        // need GPU acceleration, and disabling it avoids loading the GPU driver stack
        // (on Intel iGPUs that pulls in ~13 DLLs, ~80 extra threads and ~160MB of RAM).
        RenderOptions.ProcessRenderMode = RenderMode.SoftwareOnly;

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

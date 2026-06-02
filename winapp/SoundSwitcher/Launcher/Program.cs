using System.Diagnostics;
using System.Runtime.InteropServices;

namespace SoundSwitcher.Launcher;

/// <summary>
/// Tiny runtime-independent bootstrapper. The real app is framework-dependent
/// (tiny + fast), so it needs the .NET 8 Desktop Runtime. This launcher verifies
/// the runtime is present, offers to install it (via winget) when it isn't, and
/// then starts the app.
/// </summary>
internal static class Program
{
    private const string AppCaption = "Sound Switcher";
    private const string WingetId = "Microsoft.DotNet.DesktopRuntime.8";
    private const string DownloadPage =
        "https://dotnet.microsoft.com/download/dotnet/8.0/runtime?cid=getdotnetcore";

    [STAThread]
    private static int Main()
    {
        string baseDir = AppContext.BaseDirectory;
        string appExe = Path.Combine(baseDir, "app", "SoundSwitcherApp.exe");

        if (!File.Exists(appExe))
        {
            Show($"アプリ本体が見つかりません:\n{appExe}", MB_ICONERROR);
            return 1;
        }

        if (!IsDesktopRuntimeInstalled())
        {
            int answer = Show(
                "このアプリの実行には「.NET 8 デスクトップ ランタイム (x64)」が必要です。\n\n" +
                "今すぐインストールしますか？",
                MB_YESNO | MB_ICONQUESTION);
            if (answer != IDYES)
                return 0;

            bool installed = TryInstallViaWinget() && IsDesktopRuntimeInstalled();
            if (!installed)
            {
                // Fall back to the official download page.
                OpenUrl(DownloadPage);
                Show(
                    "自動インストールを完了できませんでした。\n" +
                    "開いたページから「.NET Desktop Runtime」をインストール後、もう一度起動してください。",
                    MB_ICONINFORMATION);
                return 0;
            }
        }

        try
        {
            Process.Start(new ProcessStartInfo(appExe)
            {
                UseShellExecute = false,
                WorkingDirectory = Path.GetDirectoryName(appExe)!,
            });
        }
        catch (Exception ex)
        {
            Show($"起動に失敗しました:\n{ex.Message}", MB_ICONERROR);
            return 1;
        }
        return 0;
    }

    /// <summary>True if any 8.x Microsoft.WindowsDesktop.App runtime is installed.</summary>
    private static bool IsDesktopRuntimeInstalled()
    {
        foreach (var root in DotnetRoots())
        {
            var sharedDir = Path.Combine(root, "shared", "Microsoft.WindowsDesktop.App");
            if (!Directory.Exists(sharedDir)) continue;
            foreach (var dir in Directory.GetDirectories(sharedDir))
            {
                if (Path.GetFileName(dir).StartsWith("8.", StringComparison.Ordinal))
                    return true;
            }
        }
        return false;
    }

    private static IEnumerable<string> DotnetRoots()
    {
        var env = Environment.GetEnvironmentVariable("DOTNET_ROOT");
        if (!string.IsNullOrEmpty(env)) yield return env;

        var pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        if (!string.IsNullOrEmpty(pf)) yield return Path.Combine(pf, "dotnet");

        yield return @"C:\Program Files\dotnet";
    }

    private static bool TryInstallViaWinget()
    {
        try
        {
            var psi = new ProcessStartInfo("winget",
                $"install --id {WingetId} --source winget --silent " +
                "--accept-source-agreements --accept-package-agreements")
            {
                UseShellExecute = true, // shows winget's own progress window
            };
            var proc = Process.Start(psi);
            if (proc == null) return false;
            proc.WaitForExit();
            // winget returns 0 on success (and for "already installed").
            return proc.ExitCode == 0;
        }
        catch
        {
            return false; // winget not available
        }
    }

    private static void OpenUrl(string url)
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
        catch { /* ignore */ }
    }

    // --- user32 MessageBox ---

    private const uint MB_YESNO = 0x4;
    private const uint MB_ICONERROR = 0x10;
    private const uint MB_ICONQUESTION = 0x20;
    private const uint MB_ICONWARNING = 0x30;
    private const uint MB_ICONINFORMATION = 0x40;
    private const int IDYES = 6;

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

    private static int Show(string text, uint type) => MessageBoxW(IntPtr.Zero, text, AppCaption, type);
}

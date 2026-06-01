using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using SoundSwitcher.Services;
using SoundSwitcher.ViewModels;

namespace SoundSwitcher;

public partial class MainWindow : Window
{
    // Width of the resize border around the window edges (DIPs). Exceeds the outer
    // shadow margin (12px) so the grab zone reaches the visible card edge.
    private const double ResizeBorder = 14.0;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainViewModel();
    }

    private void Window_SourceInitialized(object? sender, EventArgs e)
    {
        // Restore saved window position and size.
        var s = WindowSettings.Load();
        if (s.Width >= MinWidth) Width = s.Width;
        if (s.Height >= MinHeight) Height = s.Height;
        if (!double.IsNaN(s.Left) && !double.IsNaN(s.Top) && IsOnScreen(s.Left, s.Top))
        {
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = s.Left;
            Top = s.Top;
        }

        // Hook WndProc to enable resizing on a borderless (WindowStyle=None +
        // AllowsTransparency) window via custom hit-testing.
        var src = (HwndSource)PresentationSource.FromVisual(this)!;
        src.AddHook(WndProc);
    }

    private void Window_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        // Persist the *restored* bounds so a maximized window saves its normal size.
        var b = RestoreBounds;
        var s = new WindowSettings
        {
            Left = b.Left,
            Top = b.Top,
            Width = b.Width,
            Height = b.Height,
        };
        s.Save();
    }

    private bool IsOnScreen(double left, double top)
    {
        var virt = new Rect(
            SystemParameters.VirtualScreenLeft,
            SystemParameters.VirtualScreenTop,
            SystemParameters.VirtualScreenWidth,
            SystemParameters.VirtualScreenHeight);
        // Require the title bar area to be visible.
        return virt.Contains(new Point(left + 40, top + 20));
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        else
            DragMove();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

    // --- Resize hit-testing ---

    private const int WM_NCHITTEST = 0x0084;
    private const int HTLEFT = 10, HTRIGHT = 11, HTTOP = 12, HTTOPLEFT = 13,
        HTTOPRIGHT = 14, HTBOTTOM = 15, HTBOTTOMLEFT = 16, HTBOTTOMRIGHT = 17;

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg != WM_NCHITTEST || WindowState == WindowState.Maximized)
            return IntPtr.Zero;

        // Screen coordinates of the cursor.
        int sx = (short)((long)lParam & 0xFFFF);
        int sy = (short)(((long)lParam >> 16) & 0xFFFF);

        var pos = PointFromScreen(new Point(sx, sy)); // -> DIP relative to window
        double w = ActualWidth, h = ActualHeight;
        double b = ResizeBorder;

        bool left = pos.X <= b;
        bool right = pos.X >= w - b;
        bool top = pos.Y <= b;
        bool bottom = pos.Y >= h - b;

        int hit = 0;
        if (top && left) hit = HTTOPLEFT;
        else if (top && right) hit = HTTOPRIGHT;
        else if (bottom && left) hit = HTBOTTOMLEFT;
        else if (bottom && right) hit = HTBOTTOMRIGHT;
        else if (left) hit = HTLEFT;
        else if (right) hit = HTRIGHT;
        else if (top) hit = HTTOP;
        else if (bottom) hit = HTBOTTOM;

        if (hit != 0)
        {
            handled = true;
            return new IntPtr(hit);
        }
        return IntPtr.Zero;
    }
}

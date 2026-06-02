using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using WinForms = System.Windows.Forms;

namespace SoundSwitcher.Services;

/// <summary>
/// Shows the app in the notification area (system tray) instead of the taskbar.
/// The tray icon reproduces the volume glyph (Segoe MDL2 Assets U+E995) that used
/// to sit in the title bar.
/// </summary>
public sealed class TrayIconService : IDisposable
{
    private const string Glyph = ""; // volume glyph (Segoe MDL2 Assets)
    private static readonly Color Accent = Color.FromArgb(0x4F, 0x8E, 0xF7);

    private readonly WinForms.NotifyIcon _icon;
    private readonly Window _window;
    private IntPtr _hIcon;
    private bool _disposed;

    public TrayIconService(Window window, Action onRefresh)
    {
        _window = window;

        _icon = new WinForms.NotifyIcon
        {
            Icon = CreateGlyphIcon(out _hIcon),
            Text = "Sound Switcher",
            Visible = true,
        };

        // Reuse the rendered glyph as the window icon (alt-tab / system menu).
        if (_hIcon != IntPtr.Zero)
        {
            try
            {
                _window.Icon = Imaging.CreateBitmapSourceFromHIcon(
                    _hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            }
            catch { }
        }

        var menu = new WinForms.ContextMenuStrip();
        menu.Items.Add("表示 / 非表示", null, (_, _) => ToggleWindow());
        menu.Items.Add(new WinForms.ToolStripSeparator());
        menu.Items.Add("更新", null, (_, _) => onRefresh());
        menu.Items.Add("終了", null, (_, _) => Exit());
        _icon.ContextMenuStrip = menu;

        // Left click toggles window visibility.
        _icon.MouseClick += (_, e) =>
        {
            if (e.Button == WinForms.MouseButtons.Left) ToggleWindow();
        };

        // Safety net: remove the icon even if the app exits without closing the window.
        Application.Current.Exit += (_, _) => Dispose();
    }

    private void Exit()
    {
        // Remove the tray icon while the message pump is still alive so the shell
        // processes NIM_DELETE immediately (otherwise a ghost icon can linger).
        Dispose();
        Application.Current.Shutdown();
    }

    private void ToggleWindow()
    {
        if (_window.IsVisible && _window.WindowState != WindowState.Minimized)
        {
            _window.Hide();
        }
        else
        {
            _window.Show();
            _window.WindowState = WindowState.Normal;
            _window.Activate();
        }
    }

    private static Icon CreateGlyphIcon(out IntPtr hIcon)
    {
        const int size = 32;
        using var bmp = new Bitmap(size, size);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            g.Clear(Color.Transparent);
            using var font = new Font("Segoe MDL2 Assets", 22f, System.Drawing.FontStyle.Regular, GraphicsUnit.Pixel);
            using var brush = new SolidBrush(Accent);
            using var sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center,
            };
            g.DrawString(Glyph, font, brush, new RectangleF(0, 0, size, size), sf);
        }

        hIcon = bmp.GetHicon();
        // Clone so the managed Icon is independent of the native handle, which we
        // destroy on dispose.
        return (Icon)Icon.FromHandle(hIcon).Clone();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _icon.Visible = false;   // sends NIM_DELETE
        _icon.Icon?.Dispose();
        _icon.Dispose();
        if (_hIcon != IntPtr.Zero)
        {
            DestroyIcon(_hIcon);
            _hIcon = IntPtr.Zero;
        }
    }

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);
}

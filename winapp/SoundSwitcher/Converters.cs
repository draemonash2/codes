using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using SoundSwitcher.Models;

namespace SoundSwitcher;

public class BoolToWeightConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
        => value is true ? FontWeights.SemiBold : FontWeights.Normal;
    public object ConvertBack(object v, Type t, object p, CultureInfo c) => throw new NotImplementedException();
}

public class BoolToAccentBrushConverter : IValueConverter
{
    private static readonly SolidColorBrush Accent = new(Color.FromRgb(0x4F, 0x8E, 0xF7));
    private static readonly SolidColorBrush Dim = new(Color.FromRgb(0x88, 0x90, 0xA4));

    public object Convert(object value, Type t, object p, CultureInfo c)
        => value is true ? Accent : Dim;
    public object ConvertBack(object v, Type t, object p, CultureInfo c) => throw new NotImplementedException();
}

public class BatteryColorConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
    {
        if (value is not int pct) return Brushes.Transparent;
        if (pct > 40) return new SolidColorBrush(Color.FromRgb(0x3D, 0xCA, 0x7A));
        if (pct > 15) return new SolidColorBrush(Color.FromRgb(0xF5, 0xA6, 0x23));
        return new SolidColorBrush(Color.FromRgb(0xE5, 0x3E, 0x3E));
    }
    public object ConvertBack(object v, Type t, object p, CultureInfo c) => throw new NotImplementedException();
}

public class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
    {
        bool b = value is true;
        if (p is string s && s == "invert") b = !b;
        return b ? Visibility.Visible : Visibility.Collapsed;
    }
    public object ConvertBack(object v, Type t, object p, CultureInfo c) => throw new NotImplementedException();
}

public class DeviceKindToIconConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
        => value is DeviceKind.Recording ? "" : ""; // Microphone : Volume
    public object ConvertBack(object v, Type t, object p, CultureInfo c) => throw new NotImplementedException();
}

/// <summary>
/// Converts battery percent (0-100) to the endpoint Point of an arc on a circle:
/// center (32,32), radius 29, starting at top (32,3), going clockwise.
/// </summary>
public class BatteryArcPointConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
    {
        int pct = value is int v ? v : 0;
        // Clamp so 100% doesn't collapse the arc (use 99.9%)
        double fraction = Math.Min(pct / 100.0, 0.999);
        double angle = fraction * 360.0 - 90.0; // start at top = -90 deg
        double rad = angle * Math.PI / 180.0;
        double x = 32 + 29 * Math.Cos(rad);
        double y = 32 + 29 * Math.Sin(rad);
        return new Point(x, y);
    }
    public object ConvertBack(object v, Type t, object p, CultureInfo c) => throw new NotImplementedException();
}

public class BatteryIsLargeArcConverter : IValueConverter
{
    public object Convert(object value, Type t, object p, CultureInfo c)
        => value is int pct && pct > 50;
    public object ConvertBack(object v, Type t, object p, CultureInfo c) => throw new NotImplementedException();
}

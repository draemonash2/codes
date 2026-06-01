using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace SoundSwitcher.Models;

public enum DeviceKind { Playback, Recording }
public enum ConnectionState { Connected, Disconnected, NotPresent }

public class AudioDevice : INotifyPropertyChanged
{
    private bool _isDefault;
    private int? _batteryPercent;
    private ConnectionState _connectionState;

    public string Id { get; init; } = "";
    public string Name { get; init; } = "";
    public DeviceKind Kind { get; init; }
    public string IconPath { get; init; } = "";

    public bool IsDefault
    {
        get => _isDefault;
        set { _isDefault = value; OnPropertyChanged(); }
    }

    public int? BatteryPercent
    {
        get => _batteryPercent;
        set { _batteryPercent = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasBattery)); OnPropertyChanged(nameof(BatteryArcLength)); }
    }

    public ConnectionState ConnectionState
    {
        get => _connectionState;
        set { _connectionState = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsConnected)); OnPropertyChanged(nameof(OpacityValue)); }
    }

    public bool HasBattery => _batteryPercent.HasValue;
    public bool IsConnected => _connectionState == ConnectionState.Connected;
    public double OpacityValue => IsConnected ? 1.0 : 0.4;

    // Arc length for battery indicator (circumference of circle r=28 is ~175.9)
    public double BatteryArcLength => HasBattery ? (_batteryPercent!.Value / 100.0) * 175.9 : 0;

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

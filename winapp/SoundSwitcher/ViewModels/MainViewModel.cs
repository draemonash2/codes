using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using System.Windows.Threading;
using NAudio.CoreAudioApi;
using SoundSwitcher.Models;
using SoundSwitcher.Services;

namespace SoundSwitcher.ViewModels;

public class RelayCommand(Action<object?> execute) : ICommand
{
    public bool CanExecute(object? parameter) => true;
    public void Execute(object? parameter) => execute(parameter);
    public event EventHandler? CanExecuteChanged;
}

public class MainViewModel : INotifyPropertyChanged
{
    private readonly AudioDeviceService _service;

    public ObservableCollection<AudioDevice> PlaybackDevices => _service.PlaybackDevices;
    public ObservableCollection<AudioDevice> RecordingDevices => _service.RecordingDevices;

    // Master volume of the default output (playback) and input (recording) endpoints.
    public EndpointVolume OutputVolume { get; } = new();
    public EndpointVolume InputVolume { get; } = new();

    // Drives the level meters. Runs only while the window is visible AND meters are on.
    private readonly DispatcherTimer _meterTimer;
    // Re-reads Bluetooth battery levels periodically while the window is visible
    // (the OS doesn't push battery changes).
    private readonly DispatcherTimer _batteryTimer;
    private bool _windowShown;
    private bool _metersVisible = WindowSettings.Load().MetersVisible; // persisted; default off

    /// <summary>Toggles the level meters (UI + the sampling timer). Off removes their
    /// entire cost, which is handy for isolating their CPU use.</summary>
    public bool MetersVisible
    {
        get => _metersVisible;
        set { if (_metersVisible == value) return; _metersVisible = value; OnPropertyChanged(); UpdateMetering(); }
    }

    public ICommand SetDefaultCommand { get; }
    public ICommand ToggleOutputMuteCommand { get; }
    public ICommand ToggleInputMuteCommand { get; }

    public MainViewModel()
    {
        _service = new AudioDeviceService();
        SetDefaultCommand = new RelayCommand(SetDefault);
        ToggleOutputMuteCommand = new RelayCommand(_ => OutputVolume.ToggleMute());
        ToggleInputMuteCommand = new RelayCommand(_ => InputVolume.ToggleMute());

        // Re-point the volume sliders whenever the default device changes.
        _service.DevicesChanged += AttachVolumes;
        AttachVolumes();

        _meterTimer = new DispatcherTimer(DispatcherPriority.Render)
        {
            Interval = TimeSpan.FromMilliseconds(100), // ~10 fps (keeps CPU low during calls)
        };
        _meterTimer.Tick += (_, _) =>
        {
            OutputVolume.SamplePeak();
            InputVolume.SamplePeak();
        };

        _batteryTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromSeconds(30),
        };
        _batteryTimer.Tick += (_, _) => _service.RefreshBatteries();
    }

    /// <summary>Called by the window when it is shown/hidden (it lives in the tray).</summary>
    public void SetWindowVisible(bool shown)
    {
        _windowShown = shown;
        UpdateMetering();

        if (shown)
        {
            _service.RefreshBatteries(); // refresh immediately on show, then every 30 s
            _batteryTimer.Start();
        }
        else
        {
            _batteryTimer.Stop();
        }
    }

    // The timer runs only when the window is shown and the meters are enabled.
    private void UpdateMetering()
    {
        if (_windowShown && _metersVisible) _meterTimer.Start();
        else _meterTimer.Stop();
    }

    private void AttachVolumes()
    {
        OutputVolume.Attach(_service.GetDefaultDevice(DataFlow.Render));
        InputVolume.Attach(_service.GetDefaultDevice(DataFlow.Capture));
    }

    private void SetDefault(object? parameter)
    {
        if (parameter is AudioDevice device)
            _service.SetDefaultDevice(device);
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

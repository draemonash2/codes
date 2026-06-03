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

    // Drives the level meters. Runs only while the window is visible (see Start/Stop).
    private readonly DispatcherTimer _meterTimer;

    public ICommand SetDefaultCommand { get; }
    public ICommand RefreshCommand { get; }

    public MainViewModel()
    {
        _service = new AudioDeviceService();
        SetDefaultCommand = new RelayCommand(SetDefault);
        RefreshCommand = new RelayCommand(_ => _service.Refresh());

        // Re-point the volume sliders whenever the default device changes.
        _service.DevicesChanged += AttachVolumes;
        AttachVolumes();

        _meterTimer = new DispatcherTimer(DispatcherPriority.Render)
        {
            Interval = TimeSpan.FromMilliseconds(33), // ~30 fps
        };
        _meterTimer.Tick += (_, _) =>
        {
            OutputVolume.SamplePeak();
            InputVolume.SamplePeak();
        };
    }

    /// <summary>Begin/stop sampling the level meters (tie to window visibility).</summary>
    public void StartMetering() => _meterTimer.Start();
    public void StopMetering() => _meterTimer.Stop();

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

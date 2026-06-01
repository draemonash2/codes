using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
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

    public ICommand SetDefaultCommand { get; }
    public ICommand RefreshCommand { get; }

    public MainViewModel()
    {
        _service = new AudioDeviceService();
        SetDefaultCommand = new RelayCommand(SetDefault);
        RefreshCommand = new RelayCommand(_ => _service.Refresh());
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

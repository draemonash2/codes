using System.ComponentModel;
using System.Runtime.CompilerServices;
using NAudio.CoreAudioApi;

namespace SoundSwitcher.Services;

/// <summary>
/// Exposes the master volume + mute of one audio endpoint as a single 0..1 value
/// where 0 means muted. Tracks whichever device is currently attached and reflects
/// external changes (hardware keys, the Windows mixer) back to the bound UI.
/// Also surfaces a live signal <see cref="Peak"/> for a level meter.
/// </summary>
public sealed class EndpointVolume : INotifyPropertyChanged, IDisposable
{
    private MMDevice? _device;
    private AudioEndpointVolume? _endpoint;
    private AudioMeterInformation? _meter;
    private bool _applying; // suppresses the feedback notification while we set the value

    public bool HasDevice => _endpoint != null;

    /// <summary>
    /// Current signal level, 0..1, for a level meter. Updated by <see cref="SamplePeak"/>.
    /// Falls back smoothly so the bar drops gradually instead of flickering.
    /// </summary>
    /// <remarks>
    /// This reads the endpoint's hardware peak meter, which requires no audio stream
    /// of our own. For a render (speaker) endpoint it reflects whatever the system is
    /// playing. For a capture (mic) endpoint it only moves while *some* app is actively
    /// capturing — Windows produces no signal otherwise, and we deliberately do not
    /// open a capture stream just to drive the meter (that would hold the microphone).
    /// </remarks>
    public double Peak { get; private set; }

    /// <summary>Read the live peak meter and update <see cref="Peak"/> (call on a UI timer).</summary>
    public void SamplePeak()
    {
        double raw = 0;
        if (_meter != null)
        {
            try { raw = _meter.MasterPeakValue; } catch { raw = 0; }
        }
        // Instant rise, smooth decay (classic peak-meter ballistics).
        double next = raw >= Peak ? raw : Peak * 0.80;
        if (next < 0.001) next = 0;
        if (Math.Abs(next - Peak) > 0.0005)
        {
            Peak = next;
            OnPropertyChanged(nameof(Peak));
        }
    }

    /// <summary>
    /// 0..1, where 0 == muted. Dragging to the bottom (0) mutes the endpoint;
    /// any value above 0 unmutes it and sets the master level.
    /// </summary>
    public double Volume
    {
        get
        {
            if (_endpoint == null) return 0;
            try { return _endpoint.Mute ? 0 : _endpoint.MasterVolumeLevelScalar; }
            catch { return 0; }
        }
        set
        {
            if (_endpoint == null) return;
            _applying = true;
            try
            {
                float v = (float)Math.Clamp(value, 0, 1);
                if (v <= 0.0001f)
                {
                    _endpoint.Mute = true;
                }
                else
                {
                    _endpoint.MasterVolumeLevelScalar = v;
                    _endpoint.Mute = false;
                }
            }
            catch { }
            finally { _applying = false; }
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsMuted));
        }
    }

    /// <summary>Whether the endpoint is currently muted.</summary>
    public bool IsMuted
    {
        get
        {
            if (_endpoint == null) return false;
            try { return _endpoint.Mute; } catch { return false; }
        }
    }

    /// <summary>Toggle mute, preserving the underlying volume level.</summary>
    public void ToggleMute()
    {
        if (_endpoint == null) return;
        _applying = true;
        try { _endpoint.Mute = !_endpoint.Mute; } catch { }
        finally { _applying = false; }
        OnPropertyChanged(nameof(Volume));
        OnPropertyChanged(nameof(IsMuted));
    }

    /// <summary>Point this controller at a new endpoint (e.g. after the default changes).</summary>
    public void Attach(MMDevice? device)
    {
        if (_endpoint != null)
        {
            try { _endpoint.OnVolumeNotification -= OnNotify; } catch { }
        }
        _device?.Dispose();

        _device = device;
        _endpoint = null;
        _meter = null;
        if (device != null)
        {
            try
            {
                _endpoint = device.AudioEndpointVolume;
                _endpoint.OnVolumeNotification += OnNotify;
            }
            catch { _endpoint = null; }
            try { _meter = device.AudioMeterInformation; }
            catch { _meter = null; }
        }

        Peak = 0;
        OnPropertyChanged(nameof(Volume));
        OnPropertyChanged(nameof(HasDevice));
        OnPropertyChanged(nameof(Peak));
        OnPropertyChanged(nameof(IsMuted));
    }

    // Fired on a COM thread when the volume/mute changes elsewhere; marshal to the UI.
    private void OnNotify(AudioVolumeNotificationData data)
    {
        if (_applying) return;
        App.Current?.Dispatcher.BeginInvoke(() =>
        {
            OnPropertyChanged(nameof(Volume));
            OnPropertyChanged(nameof(IsMuted));
        });
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    public void Dispose()
    {
        if (_endpoint != null)
        {
            try { _endpoint.OnVolumeNotification -= OnNotify; } catch { }
        }
        _device?.Dispose();
        _device = null;
        _endpoint = null;
        _meter = null;
    }
}

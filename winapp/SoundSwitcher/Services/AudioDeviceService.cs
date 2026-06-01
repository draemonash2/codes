using System.Collections.ObjectModel;
using System.Management;
using System.Runtime.InteropServices;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using SoundSwitcher.Models;

namespace SoundSwitcher.Services;

public class AudioDeviceService : IMMNotificationClient, IDisposable
{
    private readonly MMDeviceEnumerator _enumerator;
    public ObservableCollection<AudioDevice> PlaybackDevices { get; } = new();
    public ObservableCollection<AudioDevice> RecordingDevices { get; } = new();

    public event Action? DevicesChanged;

    public AudioDeviceService()
    {
        _enumerator = new MMDeviceEnumerator();
        _enumerator.RegisterEndpointNotificationCallback(this);
        Refresh();
    }

    public void Refresh()
    {
        PlaybackDevices.Clear();
        RecordingDevices.Clear();

        var defaultPlaybackId = GetDefaultDeviceId(DataFlow.Render);
        var defaultRecordingId = GetDefaultDeviceId(DataFlow.Capture);

        AddDevices(DataFlow.Render, DeviceKind.Playback, defaultPlaybackId);
        AddDevices(DataFlow.Capture, DeviceKind.Recording, defaultRecordingId);
    }

    private string GetDefaultDeviceId(DataFlow flow)
    {
        try
        {
            return _enumerator.GetDefaultAudioEndpoint(flow, Role.Multimedia)?.ID ?? "";
        }
        catch { return ""; }
    }

    private void AddDevices(DataFlow flow, DeviceKind kind, string defaultId)
    {
        var target = kind == DeviceKind.Playback ? PlaybackDevices : RecordingDevices;

        // Active = connected, Unplugged = paired BT device not connected
        var deviceStates = new[] {
            (DeviceState.Active, ConnectionState.Connected),
            (DeviceState.Unplugged, ConnectionState.Disconnected),
        };

        var seen = new HashSet<string>();
        foreach (var (state, connState) in deviceStates)
        {
            MMDeviceCollection? devices = null;
            try { devices = _enumerator.EnumerateAudioEndPoints(flow, state); }
            catch { continue; }

            foreach (var dev in devices)
            {
                string id;
                try { id = dev.ID; } catch { continue; }
                if (seen.Contains(id)) continue;
                seen.Add(id);

                var name = GetFriendlyName(dev);
                if (string.IsNullOrEmpty(name)) continue;

                var battery = connState == ConnectionState.Connected ? GetBatteryLevel(name) : null;
                target.Add(new AudioDevice
                {
                    Id = id,
                    Name = name,
                    Kind = kind,
                    IsDefault = id == defaultId,
                    BatteryPercent = battery,
                    ConnectionState = connState,
                });
            }
        }
    }

    private static string GetFriendlyName(MMDevice dev)
    {
        try { return dev.FriendlyName; } catch { }
        try { return dev.DeviceFriendlyName; } catch { }
        return "";
    }

    private static int? GetBatteryLevel(string deviceName)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                $"SELECT * FROM Win32_Battery WHERE Name LIKE '%{EscapeWql(deviceName)}%'");
            foreach (ManagementObject bat in searcher.Get())
            {
                if (bat["EstimatedChargeRemaining"] is ushort pct)
                    return pct;
            }
        }
        catch { }
        return null;
    }

    private static string EscapeWql(string s) => s.Replace("'", "\\'").Replace("\\", "\\\\");

    public void SetDefaultDevice(AudioDevice device)
    {
        PolicyConfigClient.SetDefaultEndpoint(device.Id, Role.Multimedia);
        PolicyConfigClient.SetDefaultEndpoint(device.Id, Role.Communications);
        Refresh();
    }

    // IMMNotificationClient — refresh on any device change
    void IMMNotificationClient.OnDeviceStateChanged(string deviceId, DeviceState newState)
        => App.Current.Dispatcher.BeginInvoke(Refresh);
    void IMMNotificationClient.OnDeviceAdded(string pwstrDeviceId)
        => App.Current.Dispatcher.BeginInvoke(Refresh);
    void IMMNotificationClient.OnDeviceRemoved(string deviceId)
        => App.Current.Dispatcher.BeginInvoke(Refresh);
    void IMMNotificationClient.OnDefaultDeviceChanged(DataFlow flow, Role role, string defaultDeviceId)
        => App.Current.Dispatcher.BeginInvoke(Refresh);
    void IMMNotificationClient.OnPropertyValueChanged(string pwstrDeviceId, PropertyKey key) { }

    public void Dispose()
    {
        _enumerator.UnregisterEndpointNotificationCallback(this);
        _enumerator.Dispose();
    }
}

// PolicyConfigClient: undocumented COM interface to set default audio endpoint
internal static class PolicyConfigClient
{
    [ComImport, Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
    private class CPolicyConfigClient { }

    [ComImport, Guid("f8679f50-850a-41cf-9c72-430f290290c8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPolicyConfig
    {
        [PreserveSig] int GetMixFormat(string pszDeviceName, IntPtr ppFormat);
        [PreserveSig] int GetDeviceFormat(string pszDeviceName, bool bDefault, IntPtr ppFormat);
        [PreserveSig] int ResetDeviceFormat(string pszDeviceName);
        [PreserveSig] int SetDeviceFormat(string pszDeviceName, IntPtr pEndpointFormat, IntPtr MixFormat);
        [PreserveSig] int GetProcessingPeriod(string pszDeviceName, bool bDefault, IntPtr pmftDefaultPeriod, IntPtr pmftMinimumPeriod);
        [PreserveSig] int SetProcessingPeriod(string pszDeviceName, IntPtr pmftPeriod);
        [PreserveSig] int GetShareMode(string pszDeviceName, IntPtr pMode);
        [PreserveSig] int SetShareMode(string pszDeviceName, IntPtr mode);
        [PreserveSig] int GetPropertyValue(string pszDeviceName, bool bFxStore, IntPtr key, IntPtr pv);
        [PreserveSig] int SetPropertyValue(string pszDeviceName, bool bFxStore, IntPtr key, IntPtr pv);
        [PreserveSig] int SetDefaultEndpoint(string pszDeviceName, Role role);
        [PreserveSig] int SetEndpointVisibility(string pszDeviceName, bool bVisible);
    }

    public static void SetDefaultEndpoint(string deviceId, Role role)
    {
        var policyConfig = (IPolicyConfig)new CPolicyConfigClient();
        policyConfig.SetDefaultEndpoint(deviceId, role);
    }
}

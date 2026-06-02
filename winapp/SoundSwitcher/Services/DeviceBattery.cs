using System.Runtime.InteropServices;

namespace SoundSwitcher.Services;

/// <summary>
/// Reads battery levels for Bluetooth devices via the PnP device property
/// DEVPKEY_Bluetooth_Battery, using SetupAPI. This is the property Windows itself
/// shows in Settings &gt; Bluetooth &amp; devices, and works for connected BT audio
/// gear (headsets/earbuds) that report battery over HFP/BLE.
/// </summary>
public static class DeviceBattery
{
    // Snapshot of (device friendly name, battery 0-100) for all devices that
    // currently expose a Bluetooth battery level.
    public static List<(string Name, int Battery)> Snapshot()
    {
        var result = new List<(string, int)>();

        IntPtr set = SetupDiGetClassDevs(IntPtr.Zero, null, IntPtr.Zero,
            DIGCF_PRESENT | DIGCF_ALLCLASSES);
        if (set == INVALID_HANDLE_VALUE)
            return result;

        try
        {
            var data = new SP_DEVINFO_DATA();
            data.cbSize = (uint)Marshal.SizeOf<SP_DEVINFO_DATA>();

            for (uint i = 0; SetupDiEnumDeviceInfo(set, i, ref data); i++)
            {
                if (TryGetByteProperty(set, ref data, in DEVPKEY_Bluetooth_Battery, out byte battery))
                {
                    string name = GetStringProperty(set, ref data, in DEVPKEY_Device_FriendlyName)
                                  ?? GetStringProperty(set, ref data, in DEVPKEY_NAME)
                                  ?? "";
                    if (!string.IsNullOrWhiteSpace(name))
                        result.Add((name, battery));
                }
            }
        }
        finally
        {
            SetupDiDestroyDeviceInfoList(set);
        }

        return result;
    }

    /// <summary>
    /// Finds a battery level for an audio endpoint by matching its name against the
    /// Bluetooth device names. Audio endpoint names look like
    /// "ヘッドホン (SOUNDPEATS Air4)"; the BT device name is "SOUNDPEATS Air4".
    /// </summary>
    public static int? Match(string audioDeviceName, List<(string Name, int Battery)> snapshot)
    {
        if (snapshot.Count == 0) return null;

        // Prefer the text inside the trailing parentheses, which holds the BT name.
        string key = audioDeviceName;
        int open = audioDeviceName.LastIndexOf('(');
        int close = audioDeviceName.LastIndexOf(')');
        if (open >= 0 && close > open)
            key = audioDeviceName.Substring(open + 1, close - open - 1).Trim();

        foreach (var (name, battery) in snapshot)
        {
            if (key.Contains(name, StringComparison.OrdinalIgnoreCase) ||
                name.Contains(key, StringComparison.OrdinalIgnoreCase))
                return battery;
        }
        return null;
    }

    // --- SetupAPI interop ---

    private const uint DIGCF_PRESENT = 0x02;
    private const uint DIGCF_ALLCLASSES = 0x04;
    private const uint DEVPROP_TYPE_BYTE = 0x0003;
    private const uint DEVPROP_TYPE_STRING = 0x0012;
    private static readonly IntPtr INVALID_HANDLE_VALUE = new(-1);

    [StructLayout(LayoutKind.Sequential)]
    private struct SP_DEVINFO_DATA
    {
        public uint cbSize;
        public Guid ClassGuid;
        public uint DevInst;
        public IntPtr Reserved;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DEVPROPKEY
    {
        public Guid fmtid;
        public uint pid;
    }

    // DEVPKEY_Bluetooth_Battery {104EA319-6EE2-4701-BD47-8DDBF425BBE5} 2
    private static readonly DEVPROPKEY DEVPKEY_Bluetooth_Battery = new()
    {
        fmtid = new Guid(0x104EA319, 0x6EE2, 0x4701, 0xBD, 0x47, 0x8D, 0xDB, 0xF4, 0x25, 0xBB, 0xE5),
        pid = 2,
    };

    // DEVPKEY_Device_FriendlyName {A45C254E-DF1C-4EFD-8020-67D146A850E0} 14
    private static readonly DEVPROPKEY DEVPKEY_Device_FriendlyName = new()
    {
        fmtid = new Guid(0xA45C254E, 0xDF1C, 0x4EFD, 0x80, 0x20, 0x67, 0xD1, 0x46, 0xA8, 0x50, 0xE0),
        pid = 14,
    };

    // DEVPKEY_NAME {B725F130-47EF-101A-A5F1-02608C9EEBAC} 10
    private static readonly DEVPROPKEY DEVPKEY_NAME = new()
    {
        fmtid = new Guid(0xB725F130, 0x47EF, 0x101A, 0xA5, 0xF1, 0x02, 0x60, 0x8C, 0x9E, 0xEB, 0xAC),
        pid = 10,
    };

    private static bool TryGetByteProperty(IntPtr set, ref SP_DEVINFO_DATA data,
        in DEVPROPKEY key, out byte value)
    {
        value = 0;
        var buffer = new byte[1];
        if (SetupDiGetDeviceProperty(set, ref data, in key, out uint propType,
                buffer, (uint)buffer.Length, out _, 0)
            && propType == DEVPROP_TYPE_BYTE)
        {
            value = buffer[0];
            return true;
        }
        return false;
    }

    private static string? GetStringProperty(IntPtr set, ref SP_DEVINFO_DATA data,
        in DEVPROPKEY key)
    {
        // First call to size the buffer.
        SetupDiGetDeviceProperty(set, ref data, in key, out uint propType,
            null, 0, out uint required, 0);
        if (required == 0 || propType != DEVPROP_TYPE_STRING)
            return null;

        var buffer = new byte[required];
        if (SetupDiGetDeviceProperty(set, ref data, in key, out _,
                buffer, required, out _, 0))
        {
            // Null-terminated UTF-16.
            string s = System.Text.Encoding.Unicode.GetString(buffer);
            int nul = s.IndexOf('\0');
            return nul >= 0 ? s[..nul] : s;
        }
        return null;
    }

    [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr SetupDiGetClassDevs(IntPtr classGuid, string? enumerator,
        IntPtr hwndParent, uint flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    private static extern bool SetupDiEnumDeviceInfo(IntPtr set, uint index, ref SP_DEVINFO_DATA data);

    [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool SetupDiGetDeviceProperty(IntPtr set, ref SP_DEVINFO_DATA data,
        in DEVPROPKEY key, out uint propertyType, byte[]? buffer, uint bufferSize,
        out uint requiredSize, uint flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    private static extern bool SetupDiDestroyDeviceInfoList(IntPtr set);
}

using System.Management;

namespace SoundSwitcher.Services;

public static class BluetoothService
{
    /// <summary>
    /// Attempts to retrieve battery level (0-100) for a Bluetooth device by its friendly name.
    /// Returns null if unavailable.
    /// </summary>
    public static int? GetBatteryPercent(string deviceFriendlyName)
    {
        try
        {
            // Windows stores Bluetooth battery info in the registry / WMI
            using var searcher = new ManagementObjectSearcher(
                @"\\.\root\WMI",
                "SELECT * FROM MSBatteryClass");
            foreach (ManagementObject obj in searcher.Get())
            {
                var name = obj["InstanceName"]?.ToString() ?? "";
                if (name.Contains(deviceFriendlyName, StringComparison.OrdinalIgnoreCase))
                {
                    if (obj["RemainingCapacity"] is uint cap && obj["FullChargedCapacity"] is uint full && full > 0)
                        return (int)(cap * 100 / full);
                }
            }
        }
        catch { }

        // Fallback: Win32_Battery
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT * FROM Win32_Battery");
            foreach (ManagementObject obj in searcher.Get())
            {
                var name = obj["Name"]?.ToString() ?? "";
                if (name.Contains(deviceFriendlyName, StringComparison.OrdinalIgnoreCase))
                {
                    if (obj["EstimatedChargeRemaining"] is ushort pct)
                        return pct;
                }
            }
        }
        catch { }

        return null;
    }
}

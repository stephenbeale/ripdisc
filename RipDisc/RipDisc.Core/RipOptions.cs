namespace RipDisc.Core;

public class RipOptions
{
    public string Title { get; set; } = string.Empty;
    public bool Series { get; set; }
    public int Season { get; set; }
    public int Disc { get; set; } = 1;

    /// <summary>Empty means use the config's defaultInputDrive.</summary>
    public string Drive { get; set; } = string.Empty;

    public int DriveIndex { get; set; } = -1;

    /// <summary>Empty means use the config's defaultOutputDrive.</summary>
    public string OutputDrive { get; set; } = string.Empty;

    public bool Queue { get; set; }
    public bool Bluray { get; set; }
}

public static class DriveLetter
{
    // Same rule as Get-NormalizedDriveLetter in rip-disc.ps1: "F", "F:", "F:\", "F::" and
    // " F " all become "F:" rather than "F:\:" / "F:::".
    public static string Normalize(string drive)
    {
        var trimmed = drive.Trim().TrimEnd('\\').TrimEnd(':');
        if (trimmed.Length == 0)
            throw new ArgumentException("Drive letter is empty");
        return trimmed + ":";
    }
}

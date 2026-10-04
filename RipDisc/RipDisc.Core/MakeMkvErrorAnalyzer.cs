namespace RipDisc.Core;

/// <summary>
/// Turns MakeMKV's output into a user-facing failure message. Pure functions, so the
/// classification order can be unit tested without a drive.
/// </summary>
public static class MakeMkvErrorAnalyzer
{
    /// <param name="driveDescription">How the drive is named in messages: "Drive index 1" or "D:".</param>
    /// <param name="driveHint">Label for an index-addressed drive, e.g. "External drive"; null for a drive letter.</param>
    public static string AnalyzeExitError(int exitCode, string output, int driveIndex, string driveLetter, string? driveHint)
    {
        // Check for the drive disconnecting (or the disc being ejected) mid-rip first - it has
        // its own unmistakable Windows error text and, unlike the checks below, means the disc
        // WAS read fine; something interrupted writing partway through, so "drive/disc not
        // found" would be actively misleading here.
        if (IsDeviceDisconnect(output))
            return "The drive disconnected (or the disc was ejected) partway through the rip - check the drive's USB/power connection and that the disc is still seated, then try again";

        if (ContainsAny(output, "Failed to open disc", "no disc", "can't find", "invalid drive"))
        {
            return driveIndex >= 0
                ? $"Drive not found: Drive index {driveIndex} does not exist or is not accessible"
                : $"Drive not found: {driveLetter} - verify the drive letter is correct";
        }

        if (ContainsAny(output, "no media", "medium not present", "drive is empty", "no disc in drive", "insert a disc"))
        {
            return driveIndex >= 0
                ? $"Drive is empty ({driveHint}) - please insert a disc"
                : $"Drive {driveLetter} is empty - please insert a disc";
        }

        if (ContainsAny(output, "can't access", "read error", "cannot read", "failed to read"))
            return "No disc detected in drive - the disc may be damaged or unreadable";

        return $"MakeMKV exited with code {exitCode}";
    }

    public static string AnalyzeNoFilesError(string output)
    {
        // Device-disconnect is checked first and specifically excludes the generic "no valid
        // title" match below: MakeMKV always prints a "X titles saved, Y failed" summary at the
        // end of a rip, so "0 titles saved" appears whenever EVERY title fails to save for ANY
        // reason - a disconnected drive, a full disk, permissions, anything - not just "no disc
        // was found". Matching that bare "0 titles" substring (as this used to) misreported a
        // drive that disconnected mid-save - which MakeMKV had already read titles from moments
        // earlier - as if no disc were present at all.
        if (IsDeviceDisconnect(output))
            return "The drive disconnected (or the disc was ejected) while MakeMKV was saving titles - the disc itself was read fine, but nothing could be written. Check the drive's USB/power connection and try again.";

        if (ContainsAny(output, "no valid title", "no titles found"))
            return "No disc detected in drive - MakeMKV could not find any valid titles";

        if (ContainsAny(output, "copy protection", "protected"))
            return "Disc may be copy-protected or encrypted - MakeMKV could not extract titles";

        return "No MKV files were created - check if disc is readable and contains valid content";
    }

    private static bool IsDeviceDisconnect(string output) =>
        ContainsAny(output, "STATUS_DEVICE_NOT_CONNECTED", "does not exist");

    private static bool ContainsAny(string output, params string[] needles) =>
        needles.Any(n => output.Contains(n, StringComparison.OrdinalIgnoreCase));
}

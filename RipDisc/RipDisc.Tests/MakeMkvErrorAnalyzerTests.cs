using RipDisc.Core;

namespace RipDisc.Tests;

public class MakeMkvErrorAnalyzerTests
{
    [Theory]
    [InlineData("Error 'OS error - STATUS_DEVICE_NOT_CONNECTED' occurred while reading '/VIDEO_TS/VTS_01_1.VOB'")]
    [InlineData("Error 'OS error - A device which does not exist was specified' occurred while reading '\\Device\\CdRom1'")]
    public void ExitError_DeviceDisconnect_WinsOverOtherMatches(string disconnectLine)
    {
        // "Failed to open disc" would otherwise classify this as drive-not-found
        var output = disconnectLine + "\nFailed to open disc";

        var message = MakeMkvErrorAnalyzer.AnalyzeExitError(1, output, -1, "D:", null);

        Assert.StartsWith("The drive disconnected", message);
    }

    [Fact]
    public void ExitError_DriveNotFound_ByLetter()
    {
        var message = MakeMkvErrorAnalyzer.AnalyzeExitError(1, "Failed to open disc", -1, "G:", null);
        Assert.Equal("Drive not found: G: - verify the drive letter is correct", message);
    }

    [Fact]
    public void ExitError_DriveNotFound_ByIndex()
    {
        var message = MakeMkvErrorAnalyzer.AnalyzeExitError(1, "invalid drive", 2, "D:", "External drive");
        Assert.Equal("Drive not found: Drive index 2 does not exist or is not accessible", message);
    }

    [Fact]
    public void ExitError_EmptyDrive_ByIndex_UsesDriveHint()
    {
        var message = MakeMkvErrorAnalyzer.AnalyzeExitError(1, "Medium not present", 1, "D:", "External drive");
        Assert.Equal("Drive is empty (External drive) - please insert a disc", message);
    }

    [Fact]
    public void ExitError_Unreadable()
    {
        var message = MakeMkvErrorAnalyzer.AnalyzeExitError(1, "read error at sector 123", -1, "D:", null);
        Assert.StartsWith("No disc detected in drive - the disc may be damaged", message);
    }

    [Fact]
    public void ExitError_Unrecognised_ReportsExitCode()
    {
        var message = MakeMkvErrorAnalyzer.AnalyzeExitError(7, "something else", -1, "D:", null);
        Assert.Equal("MakeMKV exited with code 7", message);
    }

    [Fact]
    public void NoFiles_DeviceDisconnect_IsNotReportedAsNoDisc()
    {
        // The real failure from 2026-08-26: titles were read, then every save failed
        var output = string.Join("\n",
            "Error 'OS error - STATUS_DEVICE_NOT_CONNECTED' occurred while reading '/VIDEO_TS/VTS_01_1.VOB'",
            "Copy complete. 0 titles saved, 21 failed.");

        var message = MakeMkvErrorAnalyzer.AnalyzeNoFilesError(output);

        Assert.StartsWith("The drive disconnected", message);
    }

    [Fact]
    public void NoFiles_ZeroTitlesSummaryAlone_IsNotReportedAsNoDisc()
    {
        // "0 titles saved" appears whenever every save fails, for any reason
        var message = MakeMkvErrorAnalyzer.AnalyzeNoFilesError("Copy complete. 0 titles saved, 3 failed.");

        Assert.DoesNotContain("No disc detected", message);
    }

    [Fact]
    public void NoFiles_NoValidTitles()
    {
        var message = MakeMkvErrorAnalyzer.AnalyzeNoFilesError("No valid titles found on disc");
        Assert.Equal("No disc detected in drive - MakeMKV could not find any valid titles", message);
    }

    [Fact]
    public void NoFiles_CopyProtection()
    {
        var message = MakeMkvErrorAnalyzer.AnalyzeNoFilesError("The disc is protected");
        Assert.StartsWith("Disc may be copy-protected", message);
    }
}

using RipDisc.Core;

namespace RipDisc.Tests;

public class DriveLetterTests
{
    [Theory]
    [InlineData("F", "F:")]
    [InlineData("F:", "F:")]
    [InlineData(@"F:\", "F:")]
    [InlineData("F::", "F:")]
    [InlineData(" F ", "F:")]
    [InlineData(@" F:\ ", "F:")]
    public void Normalize_GivesExactlyOneColon(string input, string expected)
    {
        Assert.Equal(expected, DriveLetter.Normalize(input));
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    [InlineData(@":\")]
    public void Normalize_RejectsEmpty(string input)
    {
        Assert.Throws<ArgumentException>(() => DriveLetter.Normalize(input));
    }
}

public class TitleValidatorTests
{
    [Theory]
    [InlineData("Fargo Season 1", "Season N")]
    [InlineData("Fargo Series 2", "Series N")]
    [InlineData("Fargo Disc 2", "Disc N")]
    [InlineData("Fargo S01E01", "episode code")]
    public void Series_FlagsMisplacedMetadata(string title, string expectedFragment)
    {
        var warnings = TitleValidator.GetMisplacedMetadataWarnings(title, series: true);
        Assert.Contains(warnings, w => w.Contains(expectedFragment));
    }

    [Fact]
    public void Series_CleanTitle_HasNoWarnings()
    {
        Assert.Empty(TitleValidator.GetMisplacedMetadataWarnings("Fargo", series: true));
    }

    [Fact]
    public void Movie_IsNeverWarned()
    {
        Assert.Empty(TitleValidator.GetMisplacedMetadataWarnings("Ocean's Eleven Disc 2", series: false));
    }
}

public class RipPathsTests
{
    private static readonly RipConfig Config = new() { TempRoot = @"C:\Video", DefaultInputDrive = "D:", DefaultOutputDrive = "F" };

    [Fact]
    public void Movie_Disc1()
    {
        var paths = RipPaths.Create(new RipOptions { Title = "The Matrix" }, Config);

        Assert.Equal("D:", paths.DriveLetter);
        Assert.Equal("F:", paths.OutputDriveLetter);
        Assert.Equal(@"C:\Video\The Matrix", paths.MakeMkvOutputDir);
        Assert.Equal(@"F:\DVDs\The Matrix", paths.FinalOutputDir);
        Assert.Equal(@"F:\DVDs\The Matrix\extras", paths.ExtrasDir);
    }

    [Fact]
    public void Movie_Disc2_UsesDiscSubfolder()
    {
        var paths = RipPaths.Create(new RipOptions { Title = "The Matrix", Disc = 2 }, Config);
        Assert.Equal(@"C:\Video\The Matrix\Disc2", paths.MakeMkvOutputDir);
    }

    [Fact]
    public void Series_WithSeason()
    {
        var paths = RipPaths.Create(new RipOptions { Title = "Fargo", Series = true, Season = 1 }, Config);
        Assert.Equal(@"F:\Series\Fargo\Season 1", paths.FinalOutputDir);
    }

    [Fact]
    public void Series_WithoutSeason()
    {
        var paths = RipPaths.Create(new RipOptions { Title = "Fargo", Series = true }, Config);
        Assert.Equal(@"F:\Series\Fargo", paths.FinalOutputDir);
    }

    [Fact]
    public void ExplicitDrives_OverrideConfigDefaults()
    {
        var paths = RipPaths.Create(new RipOptions { Title = "X", Drive = @"G:\", OutputDrive = "H" }, Config);

        Assert.Equal("G:", paths.DriveLetter);
        Assert.Equal(@"H:\DVDs\X", paths.FinalOutputDir);
    }
}

public class StepTrackerTests
{
    [Fact]
    public void RemainingSteps_ExcludeCompleted()
    {
        var tracker = new StepTracker();
        tracker.SetCurrentStep(1);
        tracker.CompleteCurrentStep();
        tracker.SetCurrentStep(2);

        Assert.Equal(new[] { 2, 3, 4 }, tracker.GetRemainingSteps().Select(s => s.Number));
        Assert.Single(tracker.CompletedSteps);
    }

    [Fact]
    public void CompletingTwice_DoesNotDuplicate()
    {
        var tracker = new StepTracker();
        tracker.SetCurrentStep(1);
        tracker.CompleteCurrentStep();
        tracker.CompleteCurrentStep();

        Assert.Single(tracker.CompletedSteps);
    }
}

public class FileHelperTests
{
    [Fact]
    public void GetUniqueFilePath_AddsCounterOnClash()
    {
        using var temp = new TempDir();
        File.WriteAllText(temp.Combine("a.mp4"), "");
        File.WriteAllText(temp.Combine("a-1.mp4"), "");

        Assert.Equal(temp.Combine("a-2.mp4"), FileHelper.GetUniqueFilePath(temp.Path, "a.mp4"));
        Assert.Equal(temp.Combine("b.mp4"), FileHelper.GetUniqueFilePath(temp.Path, "b.mp4"));
    }
}

public class RipConfigTests
{
    [Fact]
    public void Parse_ReadsSampleShape()
    {
        var config = RipConfig.Parse("""
            {
              "makemkvPath": "C:\\Tools\\makemkvcon64.exe",
              "handbrakePath": "",
              "tempRoot": "D:\\Rips",
              "defaultInputDrive": "G:",
              "defaultOutputDrive": "F:",
              "driveLabels": { "0": "Internal", "1": "ASUS external" },
              "tmdbApiKey": "abc",
            }
            """);

        Assert.Equal(@"C:\Tools\makemkvcon64.exe", config.MakeMkvPath);
        Assert.Equal(@"D:\Rips", config.TempRoot);
        Assert.Equal(@"D:\Rips\logs", config.LogDirectory);
        Assert.Equal(@"D:\Rips\handbrake-queue.json", config.QueueFilePath);
        Assert.Equal("G:", config.DefaultInputDrive);
        Assert.Equal("F:", config.DefaultOutputDrive);
        Assert.Equal("ASUS external", config.GetDriveLabel(1));
        Assert.Equal("unknown drive", config.GetDriveLabel(5));
        Assert.Equal("abc", config.TmdbApiKey);
    }

    [Fact]
    public void Parse_BlankValuesFallBackToDefaults()
    {
        var config = RipConfig.Parse("""{ "tempRoot": "", "defaultInputDrive": "", "defaultOutputDrive": "" }""");

        Assert.Equal(@"C:\Video", config.TempRoot);
        Assert.Equal("D:", config.DefaultInputDrive);
        Assert.Equal("E:", config.DefaultOutputDrive);
    }

    [Fact]
    public void Load_FindsConfigInParentFolder()
    {
        using var temp = new TempDir();
        File.WriteAllText(temp.Combine(RipConfig.FileName), """{ "tempRoot": "Q:\\Rips" }""");
        var nested = Directory.CreateDirectory(temp.Combine("bin", "Debug")).FullName;

        var found = RipConfig.FindConfigFile(nested);
        var config = RipConfig.Load(found);

        Assert.Equal(temp.Combine(RipConfig.FileName), found);
        Assert.Equal(@"Q:\Rips", config.TempRoot);
        Assert.Equal(found, config.SourcePath);
    }

    [Fact]
    public void Load_MissingFile_UsesDefaultDriveLabels()
    {
        using var temp = new TempDir();
        var config = RipConfig.Load(temp.Combine("does-not-exist.json"));

        Assert.Null(config.SourcePath);
        Assert.Equal("Internal drive", config.GetDriveLabel(0));
    }
}

public class CommandLineParserTests
{
    [Fact]
    public void Parse_NormalizesDrives()
    {
        var options = CommandLineParser.Parse(new[] { "-title", "X", "-drive", @"g:\", "-outputDrive", "F::" });

        Assert.Equal("g:", options.Drive);
        Assert.Equal("F:", options.OutputDrive);
    }

    [Fact]
    public void Parse_DrivesDefaultToEmpty_SoConfigDecides()
    {
        var options = CommandLineParser.Parse(new[] { "-title", "X" });

        Assert.Equal("", options.Drive);
        Assert.Equal("", options.OutputDrive);
    }

    [Fact]
    public void Parse_QueueAndProcessQueue_AreExclusive()
    {
        Assert.Throws<ArgumentException>(() => CommandLineParser.Parse(new[] { "-title", "X", "-queue", "-processQueue" }));
    }

    [Fact]
    public void Parse_ProcessQueue_NeedsNoTitle()
    {
        Assert.True(CommandLineParser.Parse(new[] { "-processQueue" }).ProcessQueue);
    }

    [Fact]
    public void Parse_MissingTitle_Throws()
    {
        Assert.Throws<ArgumentException>(() => CommandLineParser.Parse(new[] { "-series" }));
    }
}

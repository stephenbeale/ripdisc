using RipDisc.Core;

namespace RipDisc.Tests;

/// <summary>
/// Runs the real pipeline end to end against fake MakeMKV/HandBrake and a temp folder -
/// no drive, no tools, nothing ejected or opened.
/// </summary>
public sealed class RipPipelineTests : IDisposable
{
    private const long GB = 1024L * 1024 * 1024;

    private readonly TempDir _temp = new();
    private readonly FakeRipUI _ui = new();
    private readonly FakeProcessRunner _runner = new();
    private readonly FakeRipEnvironment _environment = new();
    private readonly RipConfig _config;

    public RipPipelineTests()
    {
        _config = new RipConfig
        {
            TempRoot = _temp.Combine("Video"),
            MakeMkvPath = "makemkvcon64.exe",
            HandBrakePath = "HandBrakeCLI.exe"
        };
    }

    public void Dispose() => _temp.Dispose();

    private RipPipeline CreatePipeline(RipOptions options, string? finalDir = null)
    {
        var paths = new RipPaths
        {
            DriveLetter = "D:",
            OutputDriveLetter = "X:",
            MakeMkvOutputDir = _temp.Combine("Video", options.Title),
            FinalOutputDir = finalDir ?? _temp.Combine("Out", options.Title)
        };
        return new RipPipeline(options, _config, _ui, paths, _runner, _environment);
    }

    [Fact]
    public void Movie_FullRun_RenamesFeatureAndMovesExtras()
    {
        _runner.MkvFiles["title_t00.mkv"] = 1;  // small extra
        _runner.MkvFiles["title_t01.mkv"] = 5;  // the feature
        var pipeline = CreatePipeline(new RipOptions { Title = "The Matrix" });

        var result = pipeline.Run();

        Assert.Equal(RipResult.Success, result);
        var final = _temp.Combine("Out", "The Matrix");
        Assert.True(File.Exists(Path.Combine(final, "The Matrix-Feature.mp4")));
        Assert.Equal(5, new FileInfo(Path.Combine(final, "The Matrix-Feature.mp4")).Length);
        Assert.True(File.Exists(Path.Combine(final, "extras", "The Matrix-title_t00.mp4")));
        Assert.False(Directory.Exists(_temp.Combine("Video", "The Matrix")), "temp MKV folder should be removed");
        Assert.Equal(new[] { "D:" }, _environment.Ejected);
        Assert.Equal(new[] { final }, _environment.Opened);
        Assert.Contains("The Matrix - DONE", _ui.Statuses);
        Assert.True(File.Exists(pipeline.LogFilePath));
    }

    [Fact]
    public void Series_FullRun_PrefixesWithTitle()
    {
        _runner.MkvFiles["title_t00.mkv"] = 2;
        _runner.MkvFiles["title_t01.mkv"] = 2;
        var pipeline = CreatePipeline(new RipOptions { Title = "Fargo", Series = true, Season = 1 });

        Assert.Equal(RipResult.Success, pipeline.Run());

        var final = _temp.Combine("Out", "Fargo");
        Assert.True(File.Exists(Path.Combine(final, "Fargo-title_t00.mp4")));
        Assert.True(File.Exists(Path.Combine(final, "Fargo-title_t01.mp4")));
    }

    [Fact]
    public void DecliningStart_RunsNothing()
    {
        _ui.ConfirmAnswers.Enqueue(false);
        var pipeline = CreatePipeline(new RipOptions { Title = "The Matrix" });

        var result = pipeline.Run();

        Assert.Equal(RipResult.Cancelled, result);
        Assert.Empty(_runner.Calls);
        Assert.Empty(_environment.Ejected);
    }

    [Fact]
    public void MisplacedSeriesMetadata_AsksFirst_AndDefaultsToNo()
    {
        // No queued answer, so the fake returns the prompt's default
        var pipeline = CreatePipeline(new RipOptions { Title = "Fargo Season 1", Series = true });

        var result = pipeline.Run();

        Assert.Equal(RipResult.Cancelled, result);
        Assert.Equal(new[] { "Continue with this title?" }, _ui.ConfirmPrompts);
        Assert.Empty(_runner.Calls);
    }

    [Fact]
    public void ExistingMkvFolder_ChooseSuffixed_RipsIntoNewFolder()
    {
        var existing = _temp.Combine("Video", "The Matrix");
        Directory.CreateDirectory(existing);
        File.WriteAllText(Path.Combine(existing, "old.mkv"), "keep me");
        _ui.ChooseAnswers.Enqueue(1);
        _runner.MkvFiles["title_t00.mkv"] = 3;
        var pipeline = CreatePipeline(new RipOptions { Title = "The Matrix", Queue = true });

        Assert.Equal(RipResult.Success, pipeline.Run());

        Assert.Equal("keep me", File.ReadAllText(Path.Combine(existing, "old.mkv")));
        Assert.Equal(existing + "-1", pipeline.MakeMkvOutputDir);
        Assert.Contains(_runner.Calls, c => c.Arguments.Contains(existing + "-1"));
    }

    [Fact]
    public void ExistingMkvFolder_CancelledChoice_LeavesFilesAndRunsNothing()
    {
        var existing = _temp.Combine("Video", "The Matrix");
        Directory.CreateDirectory(existing);
        File.WriteAllText(Path.Combine(existing, "old.mkv"), "keep me");
        _ui.ChooseAnswers.Enqueue(-1);
        var pipeline = CreatePipeline(new RipOptions { Title = "The Matrix" });

        Assert.Equal(RipResult.Cancelled, pipeline.Run());

        Assert.True(File.Exists(Path.Combine(existing, "old.mkv")));
        Assert.Empty(_runner.Calls);
    }

    [Fact]
    public void MakeMkvFailure_ReportsAnalysedMessage_AndDoesNotEject()
    {
        _runner.MakeMkvExitCode = 1;
        _runner.MakeMkvOutput = "Error 'OS error - STATUS_DEVICE_NOT_CONNECTED' occurred";
        var pipeline = CreatePipeline(new RipOptions { Title = "The Matrix" });

        Assert.Equal(RipResult.Failed, pipeline.Run());

        Assert.True(_ui.Wrote("Message: The drive disconnected"));
        Assert.Empty(_environment.Ejected);
        Assert.Contains("The Matrix - ERROR", _ui.Statuses);
    }

    [Fact]
    public void Bluray_SubtitleFailure_RetriesWithoutSubtitles()
    {
        _runner.MkvFiles["title_t00.mkv"] = 4;
        _runner.HandBrakeExitCode = args => args.Contains("--all-subtitles") ? 3 : 0;
        var pipeline = CreatePipeline(new RipOptions { Title = "Inception", Bluray = true });

        Assert.Equal(RipResult.Success, pipeline.Run());

        var handBrakeCalls = _runner.Calls.Where(c => c.FileName == "HandBrakeCLI.exe").ToList();
        Assert.Equal(2, handBrakeCalls.Count);
        Assert.DoesNotContain("--all-subtitles", handBrakeCalls[1].Arguments);
    }

    [Fact]
    public void OutputDriveNotReady_FailsAtEncodeStep_KeepingMkvs()
    {
        _runner.MkvFiles["title_t00.mkv"] = 4;
        var pipeline = CreatePipeline(new RipOptions { Title = "The Matrix" });
        _environment.DriveReady = false;

        Assert.Equal(RipResult.Failed, pipeline.Run());

        Assert.True(_ui.Wrote("Error at: STEP 2/4: HandBrake encoding"));
        Assert.True(File.Exists(_temp.Combine("Video", "The Matrix", "title_t00.mkv")));
    }

    [Fact]
    public void Cancellation_StopsWithCancelledResult()
    {
        _runner.MkvFiles["title_t00.mkv"] = 4;
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        var pipeline = CreatePipeline(new RipOptions { Title = "The Matrix" });

        Assert.Equal(RipResult.Cancelled, pipeline.Run(cts.Token));
        Assert.Empty(_environment.Ejected);
    }

    [Fact]
    public void QueueMode_WritesQueueEntry_AndSkipsEncoding()
    {
        _runner.MkvFiles["title_t00.mkv"] = 4;
        var pipeline = CreatePipeline(new RipOptions { Title = "The Matrix", Disc = 1, Queue = true });

        Assert.Equal(RipResult.Success, pipeline.Run());

        var queue = new RipQueue(_config.QueueFilePath).Read();
        var entry = Assert.Single(queue);
        Assert.Equal("The Matrix", entry.Title);
        Assert.Equal(pipeline.MakeMkvOutputDir, entry.MakeMkvOutputDir);
        Assert.DoesNotContain(_runner.Calls, c => c.FileName == "HandBrakeCLI.exe");
    }
}

using RipDisc.Core;

namespace RipDisc.Tests;

public sealed class QueueProcessorTests : IDisposable
{
    private readonly TempDir _temp = new();
    private readonly FakeRipUI _ui = new();
    private readonly FakeProcessRunner _runner = new();
    private readonly FakeRipEnvironment _environment = new();
    private readonly RipConfig _config;
    private readonly List<string> _encodedTitles = new();

    public QueueProcessorTests()
    {
        _config = new RipConfig
        {
            TempRoot = _temp.Combine("Video"),
            MakeMkvPath = "makemkvcon64.exe",
            HandBrakePath = "HandBrakeCLI.exe"
        };
    }

    public void Dispose() => _temp.Dispose();

    private QueueProcessor CreateProcessor() => new(_config, _ui, (options, entry) =>
    {
        _encodedTitles.Add(options.Title);
        var paths = new RipPaths
        {
            DriveLetter = "D:",
            OutputDriveLetter = "X:",
            MakeMkvOutputDir = entry.MakeMkvOutputDir,
            FinalOutputDir = _temp.Combine("Out", options.Title)
        };
        return new RipPipeline(options, _config, _ui, paths, _runner, _environment);
    });

    private QueueEntry Enqueue(string title, int minutesAgo)
    {
        var mkvDir = _temp.Combine("Video", title);
        Directory.CreateDirectory(mkvDir);
        File.WriteAllBytes(Path.Combine(mkvDir, "title_t00.mkv"), new byte[10]);

        var entry = new QueueEntry
        {
            Title = title,
            MakeMkvOutputDir = mkvDir,
            QueuedAt = DateTime.Now.AddMinutes(-minutesAgo)
        };
        new RipQueue(_config.QueueFilePath).Append(entry);
        return entry;
    }

    [Fact]
    public void ProcessesEachJobOnce_AndDeletesTheQueue()
    {
        Enqueue("Alpha", 2);
        Enqueue("Beta", 1);

        var result = CreateProcessor().ProcessAll();

        Assert.Equal(RipResult.Success, result);
        // Before the fix the finished job was merged straight back in from the file and
        // re-encoded forever
        Assert.Equal(new[] { "Alpha", "Beta" }, _encodedTitles);
        Assert.False(File.Exists(_config.QueueFilePath));
    }

    [Fact]
    public void FailedJob_StopsAndKeepsRemainingJobs()
    {
        Enqueue("Alpha", 2);
        Enqueue("Beta", 1);
        _runner.HandBrakeExitCode = _ => 1;

        var result = CreateProcessor().ProcessAll();

        Assert.Equal(RipResult.Failed, result);
        Assert.Equal(new[] { "Alpha" }, _encodedTitles);
        Assert.Equal(new[] { "Alpha", "Beta" }, new RipQueue(_config.QueueFilePath).Read().Select(e => e.Title));
    }

    [Fact]
    public void NoQueueFile_Fails()
    {
        Assert.Equal(RipResult.Failed, CreateProcessor().ProcessAll());
        Assert.True(_ui.Wrote("No queue file found."));
    }
}

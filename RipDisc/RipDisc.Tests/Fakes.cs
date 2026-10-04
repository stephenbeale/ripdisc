using RipDisc.Core;

namespace RipDisc.Tests;

/// <summary>Records output and answers prompts from queues set up by the test.</summary>
internal class FakeRipUI : IRipUI
{
    public List<(string Message, MessageKind Kind)> Messages { get; } = new();
    public List<string> ProcessOutput { get; } = new();
    public List<string> Statuses { get; } = new();
    public List<string> ConfirmPrompts { get; } = new();
    public List<string> ChoosePrompts { get; } = new();
    public Queue<bool> ConfirmAnswers { get; } = new();
    public Queue<int> ChooseAnswers { get; } = new();

    public void Write(string message, MessageKind kind = MessageKind.Info) => Messages.Add((message, kind));
    public void WriteProcessOutput(string line) => ProcessOutput.Add(line);
    public void SetStatus(string status) => Statuses.Add(status);

    public bool Confirm(string prompt, bool defaultAnswer)
    {
        ConfirmPrompts.Add(prompt);
        return ConfirmAnswers.Count > 0 ? ConfirmAnswers.Dequeue() : defaultAnswer;
    }

    public int Choose(string prompt, IReadOnlyList<string> options)
    {
        ChoosePrompts.Add(prompt);
        return ChooseAnswers.Count > 0 ? ChooseAnswers.Dequeue() : -1;
    }

    public bool Wrote(string fragment) => Messages.Any(m => m.Message.Contains(fragment));
}

/// <summary>
/// Stands in for MakeMKV and HandBrake: "rips" by writing fake MKVs into the output folder
/// named in the arguments, and "encodes" by writing the -o file.
/// </summary>
internal class FakeProcessRunner : IProcessRunner
{
    public List<(string FileName, string Arguments)> Calls { get; } = new();
    public Dictionary<string, long> MkvFiles { get; } = new();
    public int MakeMkvExitCode { get; set; }
    public string MakeMkvOutput { get; set; } = "Copy complete.";
    public Func<string, int>? HandBrakeExitCode { get; set; }

    public (int ExitCode, string Output) Run(string fileName, string arguments, string stepLabel,
        Action<string>? onOutput = null, TimeSpan? timeout = null, CancellationToken cancellationToken = default)
    {
        Calls.Add((fileName, arguments));
        cancellationToken.ThrowIfCancellationRequested();

        if (arguments.StartsWith("mkv "))
        {
            var outputDir = arguments.Split('"')[1];
            if (MakeMkvExitCode == 0)
            {
                foreach (var (name, size) in MkvFiles)
                    WriteFile(Path.Combine(outputDir, name), size);
            }
            onOutput?.Invoke(MakeMkvOutput);
            return (MakeMkvExitCode, MakeMkvOutput);
        }

        // HandBrake: -i "<input>" -o "<output>" ...
        var parts = arguments.Split('"');
        var input = parts[1];
        var output = parts[3];
        var exitCode = HandBrakeExitCode?.Invoke(arguments) ?? 0;
        if (exitCode == 0)
            WriteFile(output, new FileInfo(input).Length);
        return (exitCode, "");
    }

    private static void WriteFile(string path, long size)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var stream = File.Create(path);
        stream.SetLength(size);
    }
}

internal class FakeRipEnvironment : IRipEnvironment
{
    public List<string> Ejected { get; } = new();
    public List<string> Opened { get; } = new();
    public bool DriveReady { get; set; } = true;

    public void EjectDrive(string driveLetter) => Ejected.Add(driveLetter);
    public void OpenDirectory(string path) => Opened.Add(path);
    public (bool Ready, string Drive, string Message) TestDriveReady(string path) =>
        DriveReady ? (true, "X:", "Drive is ready") : (false, "X:", "Destination drive X: is not ready");
    public void WaitForFileHandles() { }
}

/// <summary>A throwaway folder under %TEMP%, deleted on dispose.</summary>
internal sealed class TempDir : IDisposable
{
    public TempDir()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ripdisc-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    public string Path { get; }

    public string Combine(params string[] parts) => System.IO.Path.Combine(new[] { Path }.Concat(parts).ToArray());

    public void Dispose()
    {
        try { Directory.Delete(Path, recursive: true); } catch { }
    }
}

using System.Diagnostics;
using System.Text;

namespace RipDisc.Core;

public interface IProcessRunner
{
    /// <summary>
    /// Runs an external tool to completion. <paramref name="onOutput"/> receives each
    /// stdout/stderr line as it arrives (on a thread-pool thread). Throws
    /// <see cref="ProcessingException"/> labelled with <paramref name="stepLabel"/> if the
    /// tool is missing or times out, and <see cref="OperationCanceledException"/> if
    /// <paramref name="cancellationToken"/> fires (the process tree is killed either way).
    /// </summary>
    (int ExitCode, string Output) Run(
        string fileName,
        string arguments,
        string stepLabel,
        Action<string>? onOutput = null,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default);
}

public class ProcessRunner : IProcessRunner
{
    // The timeout is a safety net against a hung external process - e.g. MakeMKV never
    // returning because a physical drive disconnected mid-rip (the PowerShell scripts in
    // this repo hit exactly that live). Output is drained asynchronously via the event
    // handlers while waiting, so a timed-out wait cannot deadlock on a full stdout/stderr
    // pipe the way a synchronous read would.
    public (int ExitCode, string Output) Run(
        string fileName,
        string arguments,
        string stepLabel,
        Action<string>? onOutput = null,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        var toolName = Path.GetFileName(fileName);
        if (!File.Exists(fileName))
            throw new ProcessingException(stepLabel,
                $"{toolName} not found at '{fileName}' - install it or set its path in {RipConfig.FileName}");

        var startInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        var outputBuilder = new StringBuilder();
        var outputLock = new object();

        void HandleLine(object sender, DataReceivedEventArgs e)
        {
            if (e.Data == null)
                return;
            lock (outputLock)
                outputBuilder.AppendLine(e.Data);
            onOutput?.Invoke(e.Data);
        }

        using var process = new Process { StartInfo = startInfo };
        process.OutputDataReceived += HandleLine;
        process.ErrorDataReceived += HandleLine;

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        var deadline = timeout.HasValue ? DateTime.UtcNow + timeout.Value : (DateTime?)null;
        while (!process.WaitForExit(250))
        {
            if (cancellationToken.IsCancellationRequested)
            {
                KillTree(process);
                cancellationToken.ThrowIfCancellationRequested();
            }

            if (deadline.HasValue && DateTime.UtcNow >= deadline.Value)
            {
                KillTree(process);
                throw new ProcessingException(stepLabel,
                    $"{toolName} did not respond within {(int)timeout!.Value.TotalSeconds}s and was terminated - the drive may have disconnected, or a leftover process from an earlier run may still be holding it. Check Task Manager for an orphaned {toolName} process and try again.");
            }
        }

        // The parameterless overload waits for the async output handlers to drain
        process.WaitForExit();

        lock (outputLock)
            return (process.ExitCode, outputBuilder.ToString());
    }

    private static void KillTree(Process process)
    {
        try { process.Kill(entireProcessTree: true); } catch { /* best effort - already stuck, not worth failing over */ }
    }
}

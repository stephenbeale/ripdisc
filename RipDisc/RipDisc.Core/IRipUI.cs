namespace RipDisc.Core;

public enum MessageKind
{
    Header,
    Success,
    Error,
    Warning,
    Info,
    Detail
}

/// <summary>
/// Everything the pipeline needs from a front end. The pipeline never touches the console
/// directly, so the same pipeline can drive the CLI and a GUI.
/// </summary>
/// <remarks>
/// Members may be called from a background thread - process output arrives on thread-pool
/// threads, and a GUI runs the pipeline off its UI thread - so a GUI implementation must
/// marshal to its UI thread.
/// </remarks>
public interface IRipUI
{
    void Write(string message, MessageKind kind = MessageKind.Info);

    /// <summary>A raw line of MakeMKV/HandBrake output.</summary>
    void WriteProcessOutput(string line);

    /// <summary>Short status for a window title, e.g. "The Matrix - DONE".</summary>
    void SetStatus(string status);

    /// <summary>
    /// Yes/no question. Must return false when no answer can be read (closed stdin, dialog
    /// dismissed) - an unanswered prompt must never be taken as consent.
    /// </summary>
    bool Confirm(string prompt, bool defaultAnswer);

    /// <summary>Returns the chosen option's index, or -1 if the user cancelled.</summary>
    int Choose(string prompt, IReadOnlyList<string> options);
}

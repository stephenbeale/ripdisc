namespace RipDisc.Core;

/// <summary>
/// Side effects on the machine outside the rip folders: ejecting the disc and opening
/// Explorer. Behind an interface so tests can run the pipeline without touching real drives.
/// </summary>
public interface IRipEnvironment
{
    void EjectDrive(string driveLetter);
    void OpenDirectory(string path);
    (bool Ready, string Drive, string Message) TestDriveReady(string path);

    /// <summary>Pause after encoding so file handles are released before cleanup.</summary>
    void WaitForFileHandles();
}

public class WindowsRipEnvironment : IRipEnvironment
{
    public void EjectDrive(string driveLetter) => FileHelper.EjectDrive(driveLetter);
    public void OpenDirectory(string path) => FileHelper.OpenDirectory(path);
    public (bool Ready, string Drive, string Message) TestDriveReady(string path) => FileHelper.TestDriveReady(path);
    public void WaitForFileHandles() => Thread.Sleep(3000);
}

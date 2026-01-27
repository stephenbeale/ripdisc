namespace RipDisc;

public static class FileHelper
{
    public static string GetUniqueFilePath(string destDir, string fileName)
    {
        var baseName = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);
        var targetPath = Path.Combine(destDir, fileName);

        if (!File.Exists(targetPath))
            return targetPath;

        int counter = 1;
        do
        {
            var newName = $"{baseName}-{counter}{extension}";
            targetPath = Path.Combine(destDir, newName);
            counter++;
        } while (File.Exists(targetPath));

        return targetPath;
    }

    public static (bool Ready, string Drive, string Message) TestDriveReady(string path)
    {
        try
        {
            var driveRoot = Path.GetPathRoot(path);
            if (string.IsNullOrEmpty(driveRoot))
            {
                return (false, "Unknown", $"Could not determine drive letter from path: {path}");
            }

            var driveDisplay = driveRoot.TrimEnd('\\');
            var driveInfo = new DriveInfo(driveDisplay);

            if (driveInfo.IsReady)
            {
                // Additional check: try to access the drive root
                if (Directory.Exists(driveRoot))
                {
                    return (true, driveDisplay, "Drive is ready");
                }
            }

            return (false, driveDisplay,
                $"Destination drive {driveDisplay} is not ready - please ensure the drive is connected and mounted");
        }
        catch
        {
            var driveRoot = Path.GetPathRoot(path);
            var driveDisplay = driveRoot?.TrimEnd('\\') ?? "Unknown";
            return (false, driveDisplay,
                $"Destination drive {driveDisplay} is not ready - please ensure the drive is connected and mounted");
        }
    }

    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public static void EjectDrive(string driveLetter)
    {
        try
        {
            // Use Shell COM object to eject disc
            var shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType != null)
            {
                dynamic? shell = Activator.CreateInstance(shellType);
                if (shell != null)
                {
                    dynamic? drive = shell.Namespace(17).ParseName(driveLetter);
                    drive?.InvokeVerb("Eject");
                }
            }
        }
        catch
        {
            // Silently ignore eject errors
        }
    }

    public static void OpenDirectory(string path)
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = path,
                UseShellExecute = true
            });
        }
        catch
        {
            // Silently ignore errors opening directory
        }
    }
}

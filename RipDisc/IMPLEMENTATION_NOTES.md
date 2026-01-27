# C# Implementation Notes

This document describes how the PowerShell script functionality was translated to C#.

## Architecture

The C# implementation uses a multi-class architecture for better organization and maintainability:

```
Program.cs                  → Entry point and command-line parsing coordination
CommandLineOptions.cs       → Parameter definitions and parsing logic
ConsoleHelper.cs           → Colored console output utilities
StepTracker.cs             → Processing step tracking and progress display
Logger.cs                  → Session logging functionality
FileHelper.cs              → File operations and utilities
RipDiscApplication.cs      → Main application logic (4 processing steps)
```

## Key Translation Mappings

### PowerShell → C# Equivalents

| PowerShell | C# |
|------------|-----|
| `param()` block | `CommandLineOptions` class |
| `[Parameter(Mandatory=$true)]` | Command-line parser validation |
| `Write-Host -ForegroundColor` | `ConsoleHelper.WriteXxx()` methods |
| `$script:variable` | Class fields (e.g., `private readonly`) |
| `function Name { }` | `private void Name()` methods |
| `$LASTEXITCODE` | `Process.ExitCode` |
| `& executable args` | `Process.Start()` with `ProcessStartInfo` |
| `Get-ChildItem` | `Directory.GetFiles()` / `Directory.GetDirectories()` |
| `Test-Path` | `File.Exists()` / `Directory.Exists()` |
| `New-Item -ItemType Directory` | `Directory.CreateDirectory()` |
| `Remove-Item` | `File.Delete()` / `Directory.Delete()` |
| `Move-Item` | `File.Move()` |
| `Rename-Item` | `File.Move()` |
| `Join-Path` | `Path.Combine()` |
| `[System.IO.Path]::GetFileNameWithoutExtension()` | `Path.GetFileNameWithoutExtension()` |
| `Start-Sleep -Seconds` | `Thread.Sleep(milliseconds)` |
| `start $path` (open explorer) | `Process.Start("explorer.exe", path)` |
| `$host.UI.RawUI.WindowTitle` | `Console.Title` |
| `Read-Host` | `Console.ReadLine()` |
| Shell.Application COM object | `Type.GetTypeFromProgID()` with dynamic |

## Detailed Implementation Notes

### 1. Command-Line Parsing

**PowerShell:**
```powershell
param(
    [Parameter(Mandatory=$true)]
    [string]$title,
    [switch]$Series,
    [int]$Season = 0,
    # ...
)
```

**C#:**
```csharp
public class CommandLineOptions
{
    public string Title { get; set; } = string.Empty;
    public bool Series { get; set; }
    public int Season { get; set; }
    // ...
}

public static CommandLineOptions Parse(string[] args)
{
    // Custom parser that handles -title, -series, etc.
}
```

### 2. Step Tracking

**PowerShell:**
```powershell
$script:AllSteps = @(
    @{ Number = 1; Name = "MakeMKV rip"; Description = "..." }
    # ...
)
$script:CompletedSteps = @()
```

**C#:**
```csharp
public class ProcessingStep
{
    public int Number { get; set; }
    public string Name { get; set; }
    public string Description { get; set; }
}

public class StepTracker
{
    private readonly List<ProcessingStep> _allSteps;
    private readonly List<ProcessingStep> _completedSteps;
    // ...
}
```

### 3. Console Output

**PowerShell:**
```powershell
Write-Host "Success!" -ForegroundColor Green
Write-Host "Warning!" -ForegroundColor Yellow
```

**C#:**
```csharp
ConsoleHelper.WriteSuccess("Success!");
ConsoleHelper.WriteWarning("Warning!");
```

The `ConsoleHelper` class encapsulates all color logic:
```csharp
public static void WriteSuccess(string message)
{
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine(message);
    Console.ResetColor();
}
```

### 4. Process Execution

**PowerShell:**
```powershell
$output = & $makemkvconPath mkv $discSource all $dir --minlength=120 2>&1
$exitCode = $LASTEXITCODE
```

**C#:**
```csharp
private (int ExitCode, string Output) ExecuteProcess(
    string fileName,
    string arguments,
    bool showOutput = false)
{
    var startInfo = new ProcessStartInfo
    {
        FileName = fileName,
        Arguments = arguments,
        UseShellExecute = false,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        CreateNoWindow = !showOutput
    };

    using var process = new Process { StartInfo = startInfo };

    // Capture output...
    process.Start();
    process.BeginOutputReadLine();
    process.BeginErrorReadLine();
    process.WaitForExit();

    return (process.ExitCode, outputBuilder.ToString());
}
```

### 5. File Operations

**PowerShell:**
```powershell
$files = Get-ChildItem -Path $dir -Filter "*.mkv"
foreach ($file in $files) {
    # Process file
}
```

**C#:**
```csharp
var files = Directory.GetFiles(dir, "*.mkv");
foreach (var file in files)
{
    // Process file
}
```

### 6. Error Handling

**PowerShell:**
```powershell
function Stop-WithError {
    param([string]$Step, [string]$Message)

    # Log error
    # Show error
    # Display manual steps
    # Open directory
    exit 1
}
```

**C#:**
```csharp
public class ProcessingException : Exception
{
    public string Step { get; }
    public ProcessingException(string step, string message)
        : base(message)
    {
        Step = step;
    }
}

private void StopWithError(string step, string message)
{
    // Log error
    // Show error
    // Display manual steps
    // Open directory
}
```

### 7. Disc Ejection (COM Interop)

**PowerShell:**
```powershell
$driveEject = New-Object -comObject Shell.Application
$driveEject.Namespace(17).ParseName($driveLetter).InvokeVerb("Eject")
```

**C#:**
```csharp
[SupportedOSPlatform("windows")]
public static void EjectDrive(string driveLetter)
{
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
```

### 8. Drive Readiness Check

**PowerShell:**
```powershell
function Test-DriveReady {
    param([string]$Path)

    $driveLetter = [System.IO.Path]::GetPathRoot($Path)
    $drive = Get-PSDrive -Name $driveDisplay.TrimEnd(':')
    # Check if ready
}
```

**C#:**
```csharp
public static (bool Ready, string Drive, string Message) TestDriveReady(string path)
{
    var driveRoot = Path.GetPathRoot(path);
    var driveInfo = new DriveInfo(driveDisplay);

    if (driveInfo.IsReady && Directory.Exists(driveRoot))
    {
        return (true, driveDisplay, "Drive is ready");
    }
    // Return error message
}
```

### 9. Logging

**PowerShell:**
```powershell
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] $Message"
    Add-Content -Path $script:LogFile -Value $entry
}
```

**C#:**
```csharp
public class Logger
{
    private readonly string _logFilePath;

    public void Log(string message)
    {
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var entry = $"[{timestamp}] {message}";
        File.AppendAllText(_logFilePath, entry + Environment.NewLine);
    }
}
```

## Testing Equivalence

To verify the C# version behaves identically to PowerShell:

1. **Same command-line parameters** - Both accept identical arguments
2. **Same directory structure** - Both create identical output directories
3. **Same file naming** - Both use identical naming conventions
4. **Same error messages** - Both provide identical error descriptions
5. **Same logging format** - Log files are identical in structure
6. **Same step tracking** - Both track and display steps identically
7. **Same window titles** - Both set identical window titles
8. **Same user prompts** - Both use identical confirmation prompts

## Performance Considerations

The C# version may offer slight performance improvements:

1. **Faster startup** - No PowerShell runtime initialization
2. **Better process handling** - Direct .NET Process API usage
3. **Compiled code** - Pre-compiled vs. interpreted scripts
4. **Lower memory footprint** - No PowerShell host overhead

However, for this workload (I/O-bound with MakeMKV and HandBrake), the differences are negligible.

## Platform Differences

Both versions are Windows-specific:

- **PowerShell**: Uses Windows PowerShell cmdlets (though could be adapted for PowerShell Core)
- **C#**: Uses `[SupportedOSPlatform("windows")]` attributes and targets `net8.0-windows`

Both require Windows COM objects for disc ejection functionality.

## Maintenance

When updating functionality:

1. Make the change in both versions
2. Test both versions with the same inputs
3. Verify identical output and behavior
4. Update both README files
5. Update this implementation notes document

## Known Differences

The only intentional differences are:

1. **Code organization** - C# uses multiple classes for better structure
2. **Type safety** - C# enforces compile-time type checking
3. **Error messages** - Minor wording differences in internal exceptions (user-facing messages are identical)

All user-visible functionality is identical.

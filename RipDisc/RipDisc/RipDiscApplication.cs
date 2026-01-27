using System.Diagnostics;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;

namespace RipDisc;

public class RipDiscApplication
{
    private readonly CommandLineOptions _options;
    private readonly StepTracker _stepTracker;
    private readonly Logger _logger;
    private string _lastWorkingDirectory = string.Empty;

    private string _makemkvOutputDir;
    private readonly string _finalOutputDir;
    private readonly string _extrasDir;
    private readonly string _makemkvconPath = @"C:\Program Files (x86)\MakeMKV\makemkvcon64.exe";
    private readonly string _handbrakePath = @"C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe";
    private readonly string _driveLetter;
    private readonly string _outputDriveLetter;
    private readonly string _windowTitle;
    private readonly bool _isMainFeatureDisc;

    public RipDiscApplication(CommandLineOptions options)
    {
        _options = options;
        _stepTracker = new StepTracker();
        _logger = new Logger(options.Title, options.Disc);

        // Normalize drive letters
        _driveLetter = options.Drive.EndsWith(":") ? options.Drive : $"{options.Drive}:";
        _outputDriveLetter = options.OutputDrive.EndsWith(":") ? options.OutputDrive : $"{options.OutputDrive}:";

        // Configure directories
        if (options.Disc > 1)
            _makemkvOutputDir = $@"C:\Video\{options.Title}\Disc{options.Disc}";
        else
            _makemkvOutputDir = $@"C:\Video\{options.Title}";

        if (options.Series)
        {
            var seriesBaseDir = $@"{_outputDriveLetter}\Series\{options.Title}";
            if (options.Season > 0)
            {
                var seasonFolder = $"Season {options.Season}";
                _finalOutputDir = Path.Combine(seriesBaseDir, seasonFolder);
            }
            else
            {
                _finalOutputDir = seriesBaseDir;
            }
        }
        else
        {
            _finalOutputDir = $@"{_outputDriveLetter}\DVDs\{options.Title}";
        }

        _extrasDir = Path.Combine(_finalOutputDir, "extras");
        _isMainFeatureDisc = !options.Series && options.Disc == 1;

        // Set window title
        if (options.Series)
        {
            _windowTitle = options.Title;
            if (options.Season > 0)
                _windowTitle += $" S{options.Season}";
            _windowTitle += $" Disc {options.Disc}";
        }
        else
        {
            _windowTitle = options.Title;
            if (options.Disc > 1)
                _windowTitle += "-extras";
        }
    }

    public int Run()
    {
        try
        {
            LogSessionStart();
            ShowDriveConfirmation();
            ConsoleHelper.SetWindowTitle(_windowTitle);
            ShowHeader();

            // Ensure extras directory exists for non-main feature discs
            EnsureExtrasDirectoryForNonMainDisc();

            // Execute the 4 processing steps
            Step1_MakeMKVRip();
            Step2_HandBrakeEncoding();
            Step3_OrganizeFiles();
            Step4_OpenDirectory();

            ShowCompletionSummary();
            ConsoleHelper.SetWindowTitle($"{_windowTitle} - DONE");
            return 0;
        }
        catch (ProcessingException ex)
        {
            StopWithError(ex.Step, ex.Message);
            return 1;
        }
        catch (Exception ex)
        {
            StopWithError("Unexpected error", ex.Message);
            return 1;
        }
    }

    private void LogSessionStart()
    {
        _logger.Log("========== RIP SESSION STARTED ==========");
        _logger.Log($"Title: {_options.Title}");
        _logger.Log($"Type: {(_options.Series ? "TV Series" : "Movie")}");
        _logger.Log($"Disc: {_options.Disc}{((_options.Disc > 1 && !_options.Series) ? " (Special Features)" : "")}");
        if (_options.Series && _options.Season > 0)
            _logger.Log($"Season: {_options.Season}");
        if (_options.DriveIndex >= 0)
            _logger.Log($"Drive Index: {_options.DriveIndex}");
        else
            _logger.Log($"Drive: {_driveLetter}");
        _logger.Log($"Output Drive: {_outputDriveLetter}");
        _logger.Log($"MakeMKV Output: {_makemkvOutputDir}");
        _logger.Log($"Final Output: {_finalOutputDir}");
        _logger.Log($"Log file: {_logger.LogFilePath}");
    }

    private void ShowDriveConfirmation()
    {
        var driveDescription = _options.DriveIndex >= 0
            ? GetDriveIndexDescription(_options.DriveIndex)
            : $"Drive {_driveLetter}";

        // Validate title doesn't contain metadata that should be separate parameters
        ValidateTitle();

        Console.WriteLine();
        ConsoleHelper.WriteSeparator();
        ConsoleHelper.WriteInfo($"Ready to rip: {_options.Title}");

        if (_options.Series)
        {
            if (_options.Season > 0)
            {
                var seasonTag = $"S{_options.Season:D2}";
                ConsoleHelper.WriteInfo($"Type: TV Series - Season {_options.Season} ({seasonTag}), Disc {_options.Disc}");
            }
            else
            {
                ConsoleHelper.WriteInfo($"Type: TV Series - Disc {_options.Disc} (no season folder)");
            }
        }
        else
        {
            var discType = _options.Disc == 1 ? "Main Feature" : "Special Features";
            ConsoleHelper.WriteInfo($"Type: Movie - {discType} (Disc {_options.Disc})");
        }

        ConsoleHelper.WriteWarning($"Using: {driveDescription}");
        ConsoleHelper.WriteWarning($"Output Drive: {_options.OutputDrive}");
        ConsoleHelper.WriteSeparator();

        ConsoleHelper.SetWindowTitle("rip-disc - INPUT");
        ConsoleHelper.ReadInput("Press Enter to continue, or Ctrl+C to abort");
    }

    private void ValidateTitle()
    {
        var warnings = new List<string>();

        if (_options.Series)
        {
            if (Regex.IsMatch(_options.Title, @"(?i)\bseries\s*\d"))
                warnings.Add("Contains 'Series N' - use -Season parameter instead");
            if (Regex.IsMatch(_options.Title, @"(?i)\bseason\s*\d"))
                warnings.Add("Contains 'Season N' - use -Season parameter instead");
            if (Regex.IsMatch(_options.Title, @"(?i)\bdisc\s*\d"))
                warnings.Add("Contains 'Disc N' - use -Disc parameter instead");
            if (Regex.IsMatch(_options.Title, @"(?i)\bS\d{1,2}E\d"))
                warnings.Add("Contains episode code (e.g. S01E01) - use -Series -Season instead");
        }

        if (warnings.Count == 0)
            return;

        Console.WriteLine();
        ConsoleHelper.WriteSeparator();
        ConsoleHelper.WriteError("WARNING: Title may contain misplaced metadata");
        ConsoleHelper.WriteSeparator();
        ConsoleHelper.WriteWarning($"Title: \"{_options.Title}\"");
        foreach (var w in warnings)
            ConsoleHelper.WriteWarning($"  ! {w}");

        Console.WriteLine();
        ConsoleHelper.WriteHeader("Expected usage:");
        ConsoleHelper.WriteInfo("  RipDisc -title \"Fargo\" -series -season 1 -disc 2");
        Console.WriteLine();

        var response = ConsoleHelper.ReadInput("Continue with this title? (y/N): ");
        if (!string.Equals(response, "y", StringComparison.OrdinalIgnoreCase))
        {
            ConsoleHelper.WriteWarning("Aborted. Please re-run with correct parameters.");
            Environment.Exit(0);
        }
    }

    private string GetDriveIndexDescription(int driveIndex)
    {
        var hint = driveIndex switch
        {
            0 => "D: internal",
            1 => "G: ASUS external",
            _ => "unknown drive"
        };
        return $"Drive Index {driveIndex} ({hint})";
    }

    private void ShowHeader()
    {
        var contentType = _options.Series ? "TV Series" : "Movie";

        Console.WriteLine();
        ConsoleHelper.WriteSeparator();
        ConsoleHelper.WriteHeader("DVD/Blu-ray Ripping & Encoding Script");
        ConsoleHelper.WriteSeparator();
        ConsoleHelper.WriteInfo($"Title: {_options.Title}");
        ConsoleHelper.WriteInfo($"Type: {contentType}");

        if (_options.DriveIndex >= 0)
        {
            var driveHint = GetDriveHint(_options.DriveIndex);
            ConsoleHelper.WriteInfo($"Drive Index: {_options.DriveIndex} ({driveHint})");
        }
        else
        {
            ConsoleHelper.WriteInfo($"Drive: {_driveLetter}");
        }

        ConsoleHelper.WriteInfo($"Output Drive: {_outputDriveLetter}");

        if (_options.Series)
        {
            if (_options.Season > 0)
            {
                var seasonTag = $"S{_options.Season:D2}";
                ConsoleHelper.WriteInfo($"Season: {_options.Season} ({seasonTag})");
            }
            else
            {
                ConsoleHelper.WriteInfo("Season: (none - no season folder)");
            }
            ConsoleHelper.WriteInfo($"Disc: {_options.Disc}");
        }
        else
        {
            var discSuffix = _options.Disc > 1 ? " (Special Features)" : "";
            ConsoleHelper.WriteInfo($"Disc: {_options.Disc}{discSuffix}");
        }

        ConsoleHelper.WriteInfo($"MakeMKV Output: {_makemkvOutputDir}");
        ConsoleHelper.WriteInfo($"Final Output: {_finalOutputDir}");
        ConsoleHelper.WriteInfo($"Log file: {_logger.LogFilePath}");
        ConsoleHelper.WriteSeparator();
        Console.WriteLine();
    }

    private string GetDriveHint(int driveIndex)
    {
        return driveIndex switch
        {
            0 => "D: internal",
            1 => "G: ASUS external",
            _ => "unknown"
        };
    }

    private void EnsureExtrasDirectoryForNonMainDisc()
    {
        if (!_isMainFeatureDisc && !_options.Series)
        {
            var driveCheck = FileHelper.TestDriveReady(_finalOutputDir);
            if (!driveCheck.Ready)
            {
                ConsoleHelper.WriteError($"\nERROR: {driveCheck.Message}");
                Environment.Exit(1);
            }

            Directory.CreateDirectory(_finalOutputDir);
            Directory.CreateDirectory(_extrasDir);
        }
    }

    private void Step1_MakeMKVRip()
    {
        _stepTracker.SetCurrentStep(1);
        _lastWorkingDirectory = _makemkvOutputDir;
        _logger.Log("STEP 1/4: Starting MakeMKV rip...");
        ConsoleHelper.WriteSuccess("[STEP 1/4] Starting MakeMKV rip...");

        // Determine disc source
        string discSource;
        if (_options.DriveIndex >= 0)
        {
            discSource = $"disc:{_options.DriveIndex}";
            ConsoleHelper.WriteSuccess($"Using drive index: {_options.DriveIndex} (bypasses drive enumeration)");
        }
        else
        {
            discSource = $"dev:{_driveLetter}";
            ConsoleHelper.WriteWarning($"Using drive: {_driveLetter} (may enumerate other drives)");
            ConsoleHelper.WriteGray("Tip: Use -DriveIndex to bypass drive enumeration");
        }

        // Handle existing directory
        ConsoleHelper.WriteWarning($"Creating directory: {_makemkvOutputDir}");
        HandleExistingMakeMKVDirectory();

        // Execute MakeMKV
        ConsoleHelper.WriteWarning("\nExecuting MakeMKV command...");
        ConsoleHelper.WriteGray($"Command: makemkvcon mkv {discSource} all \"{_makemkvOutputDir}\" --minlength=120");
        _logger.Log($"MakeMKV command: makemkvcon mkv {discSource} all \"{_makemkvOutputDir}\" --minlength=120");

        var (exitCode, output) = ExecuteProcess(_makemkvconPath,
            $"mkv {discSource} all \"{_makemkvOutputDir}\" --minlength=120", showOutput: true);

        // Check for errors
        if (exitCode != 0)
        {
            var errorMessage = AnalyzeMakeMKVError(exitCode, output);
            throw new ProcessingException("STEP 1/4: MakeMKV rip", errorMessage);
        }

        // Verify files were created
        var rippedFiles = Directory.Exists(_makemkvOutputDir)
            ? Directory.GetFiles(_makemkvOutputDir, "*.mkv")
            : Array.Empty<string>();

        if (rippedFiles.Length == 0)
        {
            var errorMessage = AnalyzeMakeMKVNoFilesError(output);
            throw new ProcessingException("STEP 1/4: MakeMKV rip", errorMessage);
        }

        // Show success
        Console.WriteLine();
        ConsoleHelper.WriteSuccess("MakeMKV rip complete!");
        ConsoleHelper.WriteInfo($"Files ripped: {rippedFiles.Length}");
        _logger.Log($"STEP 1/4: MakeMKV rip complete - {rippedFiles.Length} file(s)");

        foreach (var file in rippedFiles)
        {
            var fileInfo = new FileInfo(file);
            var sizeGB = Math.Round(fileInfo.Length / (1024.0 * 1024.0 * 1024.0), 2);
            ConsoleHelper.WriteGray($"  - {fileInfo.Name} ({sizeGB} GB)");
            _logger.Log($"  Ripped: {fileInfo.Name} ({sizeGB} GB)");
        }

        _stepTracker.CompleteCurrentStep();

        // Eject disc
        ConsoleHelper.WriteWarning($"\nEjecting disc from drive {_driveLetter}...");
        FileHelper.EjectDrive(_driveLetter);
        ConsoleHelper.WriteSuccess("Disc ejected successfully");
        _logger.Log($"Disc ejected from drive {_driveLetter}");
    }

    private void HandleExistingMakeMKVDirectory()
    {
        if (!Directory.Exists(_makemkvOutputDir))
        {
            Directory.CreateDirectory(_makemkvOutputDir);
            ConsoleHelper.WriteSuccess("Directory created successfully");
            return;
        }

        var existingFiles = Directory.GetFiles(_makemkvOutputDir);
        if (existingFiles.Length == 0)
        {
            ConsoleHelper.WriteGray("Directory exists (empty)");
            return;
        }

        // Directory exists with files
        ConsoleHelper.WriteWarning($"\nWARNING: Directory already exists with {existingFiles.Length} file(s):");
        ConsoleHelper.WriteInfo($"  {_makemkvOutputDir}");

        foreach (var file in existingFiles)
        {
            var fileInfo = new FileInfo(file);
            var sizeGB = Math.Round(fileInfo.Length / (1024.0 * 1024.0 * 1024.0), 2);
            ConsoleHelper.WriteGray($"  - {fileInfo.Name} ({sizeGB} GB)");
        }

        // Find next available suffix
        int suffix = 1;
        string suffixedDir;
        do
        {
            suffixedDir = $"{_makemkvOutputDir}-{suffix}";
            suffix++;
        } while (Directory.Exists(suffixedDir));

        ConsoleHelper.WriteHeader("\nChoose an option:");
        ConsoleHelper.WriteWarning("  [1] Delete existing files and reuse directory");
        ConsoleHelper.WriteWarning($"  [2] Use suffixed directory: {suffixedDir}");

        string? choice = null;
        while (choice != "1" && choice != "2")
        {
            choice = ConsoleHelper.ReadInput("Enter 1 or 2: ");
            if (choice != "1" && choice != "2")
                ConsoleHelper.WriteError("Invalid choice. Please enter 1 or 2.");
        }

        if (choice == "1")
        {
            ConsoleHelper.WriteWarning("Deleting existing files...");
            foreach (var file in existingFiles)
                File.Delete(file);
            ConsoleHelper.WriteSuccess($"Deleted {existingFiles.Length} existing file(s)");
            _logger.Log($"User chose to delete {existingFiles.Length} existing file(s) in {_makemkvOutputDir}");
        }
        else
        {
            _makemkvOutputDir = suffixedDir;

            Directory.CreateDirectory(suffixedDir);
            ConsoleHelper.WriteSuccess($"Using suffixed directory: {suffixedDir}");
            _logger.Log($"User chose suffixed directory: {suffixedDir}");
        }
    }

    private string AnalyzeMakeMKVError(int exitCode, string output)
    {
        var errorMessage = $"MakeMKV exited with code {exitCode}";

        // Check for drive not found
        if (output.Contains("Failed to open disc", StringComparison.OrdinalIgnoreCase) ||
            output.Contains("no disc", StringComparison.OrdinalIgnoreCase) ||
            output.Contains("can't find", StringComparison.OrdinalIgnoreCase) ||
            output.Contains("invalid drive", StringComparison.OrdinalIgnoreCase))
        {
            errorMessage = _options.DriveIndex >= 0
                ? $"Drive not found: Drive index {_options.DriveIndex} does not exist or is not accessible"
                : $"Drive not found: {_driveLetter} - verify the drive letter is correct";
            ConsoleHelper.WriteError($"\nERROR: {errorMessage}");
        }
        // Check for empty drive
        else if (output.Contains("no media", StringComparison.OrdinalIgnoreCase) ||
                 output.Contains("medium not present", StringComparison.OrdinalIgnoreCase) ||
                 output.Contains("drive is empty", StringComparison.OrdinalIgnoreCase) ||
                 output.Contains("no disc in drive", StringComparison.OrdinalIgnoreCase) ||
                 output.Contains("insert a disc", StringComparison.OrdinalIgnoreCase))
        {
            if (_options.DriveIndex >= 0)
            {
                var driveHintMsg = GetDriveHint(_options.DriveIndex);
                errorMessage = $"Drive is empty ({driveHintMsg}) - please insert a disc";
            }
            else
            {
                errorMessage = $"Drive {_driveLetter} is empty - please insert a disc";
            }
            ConsoleHelper.WriteError($"\nERROR: {errorMessage}");
        }
        // Check for disc not readable
        else if (output.Contains("can't access", StringComparison.OrdinalIgnoreCase) ||
                 output.Contains("read error", StringComparison.OrdinalIgnoreCase) ||
                 output.Contains("cannot read", StringComparison.OrdinalIgnoreCase) ||
                 output.Contains("failed to read", StringComparison.OrdinalIgnoreCase))
        {
            errorMessage = "No disc detected in drive - the disc may be damaged or unreadable";
            ConsoleHelper.WriteError($"\nERROR: {errorMessage}");
        }

        return errorMessage;
    }

    private string AnalyzeMakeMKVNoFilesError(string output)
    {
        var errorMessage = "No MKV files were created";

        if (output.Contains("no valid", StringComparison.OrdinalIgnoreCase) ||
            output.Contains("0 titles", StringComparison.OrdinalIgnoreCase))
        {
            errorMessage = "No disc detected in drive - MakeMKV could not find any valid titles";
        }
        else if (output.Contains("copy protection", StringComparison.OrdinalIgnoreCase) ||
                 output.Contains("protected", StringComparison.OrdinalIgnoreCase))
        {
            errorMessage = "Disc may be copy-protected or encrypted - MakeMKV could not extract titles";
        }
        else
        {
            errorMessage = "No MKV files were created - check if disc is readable and contains valid content";
        }

        ConsoleHelper.WriteError($"\nERROR: {errorMessage}");
        return errorMessage;
    }

    private void Step2_HandBrakeEncoding()
    {
        _stepTracker.SetCurrentStep(2);
        _lastWorkingDirectory = _finalOutputDir;
        _logger.Log("STEP 2/4: Starting HandBrake encoding...");
        ConsoleHelper.WriteSuccess("\n[STEP 2/4] Starting HandBrake encoding...");

        // Check destination drive
        ConsoleHelper.WriteWarning("Checking destination drive...");
        var driveCheck = FileHelper.TestDriveReady(_finalOutputDir);
        if (!driveCheck.Ready)
            throw new ProcessingException("STEP 2/4: HandBrake encoding", driveCheck.Message);
        ConsoleHelper.WriteSuccess($"Destination drive {driveCheck.Drive} is ready");

        // Create output directory
        ConsoleHelper.WriteWarning($"Creating directory: {_finalOutputDir}");
        if (!Directory.Exists(_finalOutputDir))
        {
            Directory.CreateDirectory(_finalOutputDir);
            ConsoleHelper.WriteSuccess("Directory created successfully");
        }
        else
        {
            ConsoleHelper.WriteWarning("Directory already exists");
        }

        // Encode each MKV file
        var mkvFiles = Directory.GetFiles(_makemkvOutputDir, "*.mkv");
        int fileCount = 0;

        foreach (var mkvFile in mkvFiles)
        {
            fileCount++;
            var mkvInfo = new FileInfo(mkvFile);
            var outputFile = Path.Combine(_finalOutputDir, Path.GetFileNameWithoutExtension(mkvFile) + ".mp4");

            Console.WriteLine();
            ConsoleHelper.WriteHeader($"--- Encoding file {fileCount} of {mkvFiles.Length} ---");
            ConsoleHelper.WriteInfo($"Input:  {mkvInfo.Name}");
            ConsoleHelper.WriteInfo($"Output: {Path.GetFileNameWithoutExtension(mkvFile)}.mp4");
            ConsoleHelper.WriteInfo($"Size:   {Math.Round(mkvInfo.Length / (1024.0 * 1024.0 * 1024.0), 2)} GB");
            _logger.Log($"Encoding file {fileCount} of {mkvFiles.Length}: {mkvInfo.Name} ({Math.Round(mkvInfo.Length / (1024.0 * 1024.0 * 1024.0), 2)} GB)");

            ConsoleHelper.WriteWarning("\nExecuting HandBrake...");
            var args = $"-i \"{mkvFile}\" -o \"{outputFile}\" " +
                      "--preset \"Fast 1080p30\" " +
                      "--all-audio " +
                      "--all-subtitles " +
                      "--subtitle-burned=none " +
                      "--verbose=1";

            var (exitCode, _) = ExecuteProcess(_handbrakePath, args, showOutput: true);

            if (exitCode != 0)
                throw new ProcessingException("STEP 2/4: HandBrake encoding",
                    $"HandBrake exited with code {exitCode} while encoding {mkvInfo.Name}");

            if (!File.Exists(outputFile))
                throw new ProcessingException("STEP 2/4: HandBrake encoding",
                    $"Output file not created for {mkvInfo.Name}");

            var outputInfo = new FileInfo(outputFile);
            Console.WriteLine();
            ConsoleHelper.WriteSuccess($"Encoding complete: {mkvInfo.Name}");
            ConsoleHelper.WriteInfo($"Output size: {Math.Round(outputInfo.Length / (1024.0 * 1024.0 * 1024.0), 2)} GB");
            _logger.Log($"Encoded: {mkvInfo.Name} -> {Path.GetFileNameWithoutExtension(mkvFile)}.mp4 ({Math.Round(outputInfo.Length / (1024.0 * 1024.0 * 1024.0), 2)} GB)");
        }

        _stepTracker.CompleteCurrentStep();
        _logger.Log($"STEP 2/4: HandBrake encoding complete - {fileCount} file(s) encoded");

        // Wait for file handles to be released
        ConsoleHelper.WriteWarning("\nWaiting for file handles to be released...");
        Thread.Sleep(3000);
        ConsoleHelper.WriteSuccess("File handle wait complete");

        // Delete MakeMKV temporary directory
        ConsoleHelper.WriteWarning("\nChecking for successful encodes...");
        var encodedFiles = Directory.GetFiles(_finalOutputDir, "*.mp4");
        if (encodedFiles.Length > 0)
        {
            ConsoleHelper.WriteSuccess($"Found {encodedFiles.Length} encoded file(s)");
            ConsoleHelper.WriteWarning($"Removing temporary MakeMKV directory: {_makemkvOutputDir}");
            Directory.Delete(_makemkvOutputDir, true);
            ConsoleHelper.WriteSuccess("Temporary files removed successfully");
            _logger.Log($"Temporary MKV directory removed: {_makemkvOutputDir}");
        }
        else
        {
            ConsoleHelper.WriteError("WARNING: No encoded files found. Keeping MakeMKV directory.");
            _logger.Log("WARNING: No encoded files found - keeping MakeMKV directory");
        }
    }

    private void Step3_OrganizeFiles()
    {
        _stepTracker.SetCurrentStep(3);
        _lastWorkingDirectory = _finalOutputDir;
        _logger.Log("STEP 3/4: Organizing files...");
        ConsoleHelper.WriteSuccess("\n[STEP 3/4] Organizing files...");

        Directory.SetCurrentDirectory(_finalOutputDir);
        ConsoleHelper.WriteWarning($"Current directory: {_finalOutputDir}");

        // Delete image files
        DeleteImageFiles();

        if (_options.Series)
        {
            PrefixSeriesFiles();
        }
        else
        {
            PrefixMovieFiles();
            if (_isMainFeatureDisc)
            {
                RenameFeatureFile();
                MoveNonFeatureToExtras();
            }
            else
            {
                MoveSpecialFeaturesToExtras();
            }
        }

        _stepTracker.CompleteCurrentStep();
        _logger.Log("STEP 3/4: File organization complete");
    }

    private void DeleteImageFiles()
    {
        ConsoleHelper.WriteWarning("\nDeleting image files...");
        var imageExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp" };
        var imageFiles = Directory.GetFiles(_finalOutputDir)
            .Where(f => imageExtensions.Contains(Path.GetExtension(f).ToLower()))
            .ToArray();

        if (imageFiles.Length > 0)
        {
            ConsoleHelper.WriteInfo($"Image files to delete: {imageFiles.Length}");
            foreach (var file in imageFiles)
            {
                ConsoleHelper.WriteGray($"  - {Path.GetFileName(file)}");
                File.Delete(file);
            }
            ConsoleHelper.WriteSuccess("Image files deleted");
        }
        else
        {
            ConsoleHelper.WriteGray("No image files found");
        }
    }

    private void PrefixSeriesFiles()
    {
        ConsoleHelper.WriteWarning("\nPrefixing files with title...");
        var filesToPrefix = Directory.GetFiles(_finalOutputDir)
            .Where(f => !Path.GetFileName(f).StartsWith($"{_options.Title}-"))
            .ToArray();

        if (filesToPrefix.Length > 0)
        {
            ConsoleHelper.WriteInfo($"Files to prefix: {filesToPrefix.Length}");
            foreach (var file in filesToPrefix)
            {
                var fileName = Path.GetFileName(file);
                ConsoleHelper.WriteGray($"  - {fileName}");
                var newPath = Path.Combine(_finalOutputDir, $"{_options.Title}-{fileName}");
                File.Move(file, newPath);
            }
            ConsoleHelper.WriteSuccess("Prefixing complete");
            _logger.Log($"Prefixed {filesToPrefix.Length} file(s) with title");
        }
        else
        {
            ConsoleHelper.WriteGray("No files need prefixing");
        }
    }

    private void PrefixMovieFiles()
    {
        var dirName = new DirectoryInfo(_finalOutputDir).Name;

        if (_isMainFeatureDisc)
        {
            ConsoleHelper.WriteWarning("\nPrefixing files with directory name...");
            var filesToPrefix = Directory.GetFiles(_finalOutputDir)
                .Where(f => !Path.GetFileName(f).StartsWith($"{dirName}-"))
                .ToArray();

            if (filesToPrefix.Length > 0)
            {
                ConsoleHelper.WriteInfo($"Files to prefix: {filesToPrefix.Length}");
                foreach (var file in filesToPrefix)
                {
                    var fileName = Path.GetFileName(file);
                    ConsoleHelper.WriteGray($"  - {fileName}");
                    var newPath = Path.Combine(_finalOutputDir, $"{dirName}-{fileName}");
                    File.Move(file, newPath);
                }
                ConsoleHelper.WriteSuccess("Prefixing complete");
                _logger.Log($"Prefixed {filesToPrefix.Length} file(s) with directory name");
            }
            else
            {
                ConsoleHelper.WriteGray("No files need prefixing");
            }
        }
        else
        {
            ConsoleHelper.WriteWarning("\nPrefixing special features files...");
            var filesToPrefix = Directory.GetFiles(_finalOutputDir)
                .Where(f => !Path.GetFileName(f).StartsWith($"{dirName}-"))
                .ToArray();

            if (filesToPrefix.Length > 0)
            {
                ConsoleHelper.WriteInfo($"Files to prefix: {filesToPrefix.Length}");
                foreach (var file in filesToPrefix)
                {
                    var fileName = Path.GetFileName(file);
                    var newName = $"{dirName}-Special Features-{fileName}";
                    ConsoleHelper.WriteGray($"  - {fileName} -> {newName}");
                    var newPath = Path.Combine(_finalOutputDir, newName);
                    File.Move(file, newPath);
                }
                ConsoleHelper.WriteSuccess("Special features prefixing complete");
                _logger.Log($"Prefixed {filesToPrefix.Length} special features file(s)");
            }
            else
            {
                ConsoleHelper.WriteGray("No files need prefixing");
            }
        }
    }

    private void RenameFeatureFile()
    {
        ConsoleHelper.WriteWarning("\nChecking for Feature file...");
        var featureExists = Directory.GetFiles(_finalOutputDir)
            .Any(f => Path.GetFileName(f).Contains("-Feature."));

        if (!featureExists)
        {
            var files = Directory.GetFiles(_finalOutputDir)
                .Select(f => new FileInfo(f))
                .OrderByDescending(f => f.Length)
                .ToArray();

            if (files.Length > 0)
            {
                var largestFile = files[0];
                var dirName = new DirectoryInfo(_finalOutputDir).Name;
                var newName = $"{dirName}-Feature{largestFile.Extension}";
                var newPath = Path.Combine(_finalOutputDir, newName);

                ConsoleHelper.WriteInfo($"Largest file: {largestFile.Name} ({Math.Round(largestFile.Length / (1024.0 * 1024.0 * 1024.0), 2)} GB)");
                ConsoleHelper.WriteWarning($"Renaming to: {newName}");
                File.Move(largestFile.FullName, newPath);
                ConsoleHelper.WriteSuccess("Feature file renamed successfully");
                _logger.Log($"Feature file: {largestFile.Name} -> {newName} ({Math.Round(largestFile.Length / (1024.0 * 1024.0 * 1024.0), 2)} GB)");
            }
        }
        else
        {
            var featureFile = Directory.GetFiles(_finalOutputDir)
                .FirstOrDefault(f => Path.GetFileName(f).Contains("-Feature."));
            ConsoleHelper.WriteGray($"Feature file already exists: {Path.GetFileName(featureFile)}");
        }
    }

    private void MoveNonFeatureToExtras()
    {
        ConsoleHelper.WriteWarning("\nChecking for non-feature videos...");
        var videoExtensions = new[] { ".mp4", ".avi", ".mkv", ".mov", ".wmv" };
        var nonFeatureVideos = Directory.GetFiles(_finalOutputDir)
            .Where(f => videoExtensions.Contains(Path.GetExtension(f).ToLower()) &&
                       !Path.GetFileName(f).Contains("Feature"))
            .ToArray();

        if (nonFeatureVideos.Length > 0)
        {
            ConsoleHelper.WriteInfo($"Non-feature videos found: {nonFeatureVideos.Length}");
            foreach (var file in nonFeatureVideos)
                ConsoleHelper.WriteGray($"  - {Path.GetFileName(file)}");

            var extrasPath = Path.Combine(_finalOutputDir, "extras");
            if (!Directory.Exists(extrasPath))
            {
                ConsoleHelper.WriteWarning("Creating extras directory...");
                Directory.CreateDirectory(extrasPath);
                ConsoleHelper.WriteSuccess("Extras directory created");
            }
            else
            {
                ConsoleHelper.WriteGray("Extras directory already exists");
            }

            ConsoleHelper.WriteWarning("Moving files to extras...");
            foreach (var file in nonFeatureVideos)
            {
                var fileName = Path.GetFileName(file);
                var destPath = Path.Combine(extrasPath, fileName);
                File.Move(file, destPath);
            }
            ConsoleHelper.WriteSuccess("Files moved to extras");
            _logger.Log($"Moved {nonFeatureVideos.Length} non-feature file(s) to extras");
        }
        else
        {
            ConsoleHelper.WriteGray("No non-feature videos found");
        }
    }

    private void MoveSpecialFeaturesToExtras()
    {
        ConsoleHelper.WriteWarning("\nMoving special features to extras folder...");

        if (!Directory.Exists(_extrasDir))
        {
            ConsoleHelper.WriteWarning("Creating extras directory...");
            Directory.CreateDirectory(_extrasDir);
            ConsoleHelper.WriteSuccess("Extras directory created");
        }
        else
        {
            ConsoleHelper.WriteGray("Extras directory already exists");
        }

        var videoExtensions = new[] { ".mp4", ".avi", ".mkv", ".mov", ".wmv" };
        var videoFiles = Directory.GetFiles(_finalOutputDir)
            .Where(f => videoExtensions.Contains(Path.GetExtension(f).ToLower()) &&
                       !Path.GetFileName(f).Contains("-Feature."))
            .ToArray();

        if (videoFiles.Length > 0)
        {
            ConsoleHelper.WriteInfo($"Videos to move: {videoFiles.Length}");
            foreach (var video in videoFiles)
            {
                var fileName = Path.GetFileName(video);
                var uniquePath = FileHelper.GetUniqueFilePath(_extrasDir, fileName);
                var newName = Path.GetFileName(uniquePath);

                if (newName != fileName)
                    ConsoleHelper.WriteWarning($"  - {fileName} -> {newName} (renamed to avoid clash)");
                else
                    ConsoleHelper.WriteGray($"  - {fileName}");

                File.Move(video, uniquePath);
            }
            ConsoleHelper.WriteSuccess("Files moved to extras");
            _logger.Log($"Moved {videoFiles.Length} special features file(s) to extras");
        }
        else
        {
            ConsoleHelper.WriteGray("No video files to move");
        }
    }

    private void Step4_OpenDirectory()
    {
        _stepTracker.SetCurrentStep(4);
        _logger.Log("STEP 4/4: Opening directory...");
        ConsoleHelper.WriteSuccess("\n[STEP 4/4] Opening film directory...");
        ConsoleHelper.WriteWarning($"Opening: {_finalOutputDir}");
        FileHelper.OpenDirectory(_finalOutputDir);
        _stepTracker.CompleteCurrentStep();
    }

    private void ShowCompletionSummary()
    {
        Console.WriteLine();
        ConsoleHelper.WriteSeparator();
        ConsoleHelper.WriteSuccess("COMPLETE!");
        ConsoleHelper.WriteSeparator();

        ConsoleHelper.WriteInfo($"\nProcessed: {GetTitleSummary()}");
        ConsoleHelper.WriteInfo($"Final location: {_finalOutputDir}");

        _stepTracker.ShowStepsSummary();

        // File summary
        Console.WriteLine();
        ConsoleHelper.WriteHeader("--- FILE SUMMARY ---");
        var finalFiles = Directory.GetFiles(_finalOutputDir, "*", SearchOption.AllDirectories);
        var totalSize = finalFiles.Sum(f => new FileInfo(f).Length);
        var totalSizeGB = Math.Round(totalSize / (1024.0 * 1024.0 * 1024.0), 2);

        ConsoleHelper.WriteInfo($"  Total files: {finalFiles.Length}");
        ConsoleHelper.WriteInfo($"  Total size: {totalSizeGB} GB");
        ConsoleHelper.WriteInfo($"  Log file: {_logger.LogFilePath}");
        ConsoleHelper.WriteSeparator();
        Console.WriteLine();

        _logger.Log("========== RIP SESSION COMPLETE ==========");
        _logger.Log($"Final location: {_finalOutputDir}");
        _logger.Log($"Total files: {finalFiles.Length}");
        _logger.Log($"Total size: {totalSizeGB} GB");
        foreach (var file in finalFiles)
        {
            var fileInfo = new FileInfo(file);
            var sizeGB = Math.Round(fileInfo.Length / (1024.0 * 1024.0 * 1024.0), 2);
            _logger.Log($"  {fileInfo.Name} ({sizeGB} GB)");
        }
    }

    private string GetTitleSummary()
    {
        var contentType = _options.Series ? "TV Series" : "Movie";
        var summary = $"{contentType}: {_options.Title}";

        if (_options.Series)
        {
            if (_options.Season > 0)
                summary += $" - Season {_options.Season}, Disc {_options.Disc}";
            else
                summary += $" - Disc {_options.Disc}";
        }
        else if (_options.Disc > 1)
        {
            summary += " (Disc " + _options.Disc + " - Special Features)";
        }

        return summary;
    }

    private void StopWithError(string step, string message)
    {
        ConsoleHelper.SetWindowTitle($"{_windowTitle} - ERROR");

        // Log the error
        _logger.Log("========== ERROR ==========");
        _logger.Log($"Failed at: {step}");
        _logger.Log($"Message: {message}");

        if (_stepTracker.CompletedSteps.Count > 0)
        {
            var completed = string.Join(", ",
                _stepTracker.CompletedSteps.Select(s => $"Step {s.Number}: {s.Name}"));
            _logger.Log($"Completed steps: {completed}");
        }
        else
        {
            _logger.Log("Completed steps: (none)");
        }

        var remaining = _stepTracker.GetRemainingSteps();
        if (remaining.Count > 0)
        {
            var remainingStr = string.Join(", ",
                remaining.Select(s => $"Step {s.Number}: {s.Name}"));
            _logger.Log($"Remaining steps: {remainingStr}");
        }

        _logger.Log($"Log file: {_logger.LogFilePath}");

        // Display error
        Console.WriteLine();
        ConsoleHelper.WriteSeparator();
        ConsoleHelper.WriteError("FAILED!");
        ConsoleHelper.WriteSeparator();

        ConsoleHelper.WriteInfo($"\nProcessing: {GetTitleSummary()}");
        Console.WriteLine();
        ConsoleHelper.WriteError($"Error at: {step}");
        ConsoleHelper.WriteError($"Message: {message}");

        _stepTracker.ShowStepsSummary(showRemaining: true);

        // Determine which directory to open
        string? directoryToOpen = null;
        if (!string.IsNullOrEmpty(_lastWorkingDirectory) && Directory.Exists(_lastWorkingDirectory))
            directoryToOpen = _lastWorkingDirectory;
        else if (Directory.Exists(_makemkvOutputDir))
            directoryToOpen = _makemkvOutputDir;
        else if (Directory.Exists(_finalOutputDir))
            directoryToOpen = _finalOutputDir;

        // Show manual steps
        ShowManualSteps(remaining);

        // Open directory
        if (directoryToOpen != null)
        {
            Console.WriteLine();
            ConsoleHelper.WriteHeader("--- OPENING DIRECTORY ---");
            ConsoleHelper.WriteWarning($"Opening: {directoryToOpen}");
            ConsoleHelper.WriteGray("(This is where leftover/partial files may be located)");
            FileHelper.OpenDirectory(directoryToOpen);
        }

        ConsoleHelper.WriteWarning($"\nLog file: {_logger.LogFilePath}");
        Console.WriteLine();
        ConsoleHelper.WriteSeparator();
        ConsoleHelper.WriteError("Please complete the remaining steps manually");
        ConsoleHelper.WriteSeparator();
        Console.WriteLine();
    }

    private void ShowManualSteps(List<ProcessingStep> remainingSteps)
    {
        Console.WriteLine();
        ConsoleHelper.WriteHeader("--- MANUAL STEPS NEEDED ---");

        foreach (var step in remainingSteps)
        {
            switch (step.Number)
            {
                case 1:
                    ConsoleHelper.WriteWarning("  - Re-run MakeMKV to rip the disc");
                    break;
                case 2:
                    ConsoleHelper.WriteWarning("  - Encode MKV files with HandBrake");
                    if (Directory.Exists(_makemkvOutputDir))
                        ConsoleHelper.WriteGray($"    MKV files location: {_makemkvOutputDir}");
                    break;
                case 3:
                    ConsoleHelper.WriteWarning("  - Rename files to proper format");
                    if (_options.Series)
                    {
                        ConsoleHelper.WriteGray($"    Format: {_options.Title}-originalname.mp4");
                    }
                    else
                    {
                        if (_options.Disc == 1)
                        {
                            ConsoleHelper.WriteGray($"    Format: {_options.Title}-Feature.mp4 (largest file)");
                            ConsoleHelper.WriteGray($"    Move extras to: {_extrasDir}");
                        }
                        else
                        {
                            ConsoleHelper.WriteGray($"    Format: {_options.Title}-Special Features-originalname.mp4");
                            ConsoleHelper.WriteGray($"    Move all files to: {_extrasDir}");
                        }
                    }
                    break;
                case 4:
                    ConsoleHelper.WriteWarning("  - Open output directory to verify files");
                    break;
            }
        }
    }

    private (int ExitCode, string Output) ExecuteProcess(string fileName, string arguments, bool showOutput = false)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        var outputBuilder = new System.Text.StringBuilder();

        using var process = new Process { StartInfo = startInfo };

        process.OutputDataReceived += (sender, e) =>
        {
            if (e.Data != null)
            {
                if (showOutput)
                    Console.WriteLine(e.Data);
                outputBuilder.AppendLine(e.Data);
            }
        };
        process.ErrorDataReceived += (sender, e) =>
        {
            if (e.Data != null)
            {
                if (showOutput)
                    Console.WriteLine(e.Data);
                outputBuilder.AppendLine(e.Data);
            }
        };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();
        process.WaitForExit();

        return (process.ExitCode, outputBuilder.ToString());
    }
}

public class ProcessingException : Exception
{
    public string Step { get; }

    public ProcessingException(string step, string message) : base(message)
    {
        Step = step;
    }
}

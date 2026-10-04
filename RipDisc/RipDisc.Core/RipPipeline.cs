namespace RipDisc.Core;

public enum RipResult
{
    Success,
    Failed,
    Cancelled
}

/// <summary>
/// The rip/encode/organize/open pipeline. All user interaction goes through
/// <see cref="IRipUI"/>, so the CLI and a GUI drive the same code.
/// </summary>
public class RipPipeline
{
    private const string Step1Label = "STEP 1/4: MakeMKV rip";
    private const string Step2Label = "STEP 2/4: HandBrake encoding";

    // 4-hour safety-net timeout: generous enough to never interrupt any realistic rip (even
    // a large Blu-ray box set), but bounded enough to eventually recover from a genuinely
    // stuck MakeMKV process instead of hanging forever with no recourse.
    private static readonly TimeSpan MakeMkvTimeout = TimeSpan.FromHours(4);

    private readonly RipOptions _options;
    private readonly RipConfig _config;
    private readonly IRipUI _ui;
    private readonly IProcessRunner _runner;
    private readonly IRipEnvironment _environment;
    private readonly StepTracker _stepTracker;
    private readonly Logger _logger;
    private string _lastWorkingDirectory = string.Empty;
    private CancellationToken _cancellationToken;

    private string _makemkvOutputDir;
    private readonly string _finalOutputDir;
    private readonly string _extrasDir;
    private readonly string _driveLetter;
    private readonly string _outputDriveLetter;
    private readonly string _windowTitle;
    private readonly bool _isMainFeatureDisc;

    public RipPipeline(RipOptions options, RipConfig config, IRipUI ui)
        : this(options, config, ui, RipPaths.Create(options, config), new ProcessRunner(), new WindowsRipEnvironment())
    {
    }

    internal RipPipeline(RipOptions options, RipConfig config, IRipUI ui, RipPaths paths,
        IProcessRunner runner, IRipEnvironment environment)
    {
        _options = options;
        _config = config;
        _ui = ui;
        _runner = runner;
        _environment = environment;
        _stepTracker = new StepTracker();
        _logger = new Logger(config.LogDirectory, options.Title, options.Disc);

        _driveLetter = paths.DriveLetter;
        _outputDriveLetter = paths.OutputDriveLetter;
        _makemkvOutputDir = paths.MakeMkvOutputDir;
        _finalOutputDir = paths.FinalOutputDir;
        _extrasDir = paths.ExtrasDir;
        _isMainFeatureDisc = !options.Series && options.Disc == 1;

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

    public string MakeMkvOutputDir => _makemkvOutputDir;
    public string FinalOutputDir => _finalOutputDir;
    public string LogFilePath => _logger.LogFilePath;

    public RipResult Run(CancellationToken cancellationToken = default)
    {
        _cancellationToken = cancellationToken;
        return Execute(() =>
        {
            LogSessionStart();
            if (!ConfirmStart())
                return RipResult.Cancelled;

            _ui.SetStatus(_windowTitle);
            ShowHeader();

            // Ensure extras directory exists for non-main feature discs
            EnsureExtrasDirectoryForNonMainDisc();

            if (!Step1_MakeMKVRip())
                return RipResult.Cancelled;

            if (_options.Queue)
            {
                // Queue mode: write job to queue file instead of encoding inline
                WriteToQueue();
                _ui.SetStatus($"{_windowTitle} - QUEUED");
            }
            else
            {
                // Normal mode: encode, organize, and open
                Step2_HandBrakeEncoding();
                Step3_OrganizeFiles();
                Step4_OpenDirectory();
                ShowCompletionSummary();
                _ui.SetStatus($"{_windowTitle} - DONE");
            }
            return RipResult.Success;
        });
    }

    public void SetMakeMkvOutputDir(string path)
    {
        _makemkvOutputDir = path;
    }

    public RipResult RunFromQueue(CancellationToken cancellationToken = default)
    {
        _cancellationToken = cancellationToken;
        return Execute(() =>
        {
            _logger.Log("========== QUEUE ENCODE SESSION ==========");
            _logger.Log($"Title: {_options.Title}");
            _logger.Log($"Type: {(_options.Series ? "TV Series" : "Movie")}");
            _logger.Log($"Disc: {_options.Disc}");
            if (_options.Series && _options.Season > 0)
                _logger.Log($"Season: {_options.Season}");
            _logger.Log($"Output: {_finalOutputDir}");

            _ui.SetStatus($"Queue: {_windowTitle}");

            if (!Directory.Exists(_makemkvOutputDir) || Directory.GetFiles(_makemkvOutputDir, "*.mkv").Length == 0)
            {
                _ui.Write($"No MKV files found in {_makemkvOutputDir}", MessageKind.Error);
                _logger.Log($"ERROR: No MKV files found in {_makemkvOutputDir}");
                return RipResult.Failed;
            }

            EnsureExtrasDirectoryForNonMainDisc();

            Step2_HandBrakeEncoding();
            Step3_OrganizeFiles();
            Step4_OpenDirectory();

            ShowCompletionSummary();
            _ui.SetStatus($"Queue: {_windowTitle} - DONE");
            return RipResult.Success;
        });
    }

    private RipResult Execute(Func<RipResult> body)
    {
        try
        {
            return body();
        }
        catch (ProcessingException ex)
        {
            StopWithError(ex.Step, ex.Message);
            return RipResult.Failed;
        }
        catch (OperationCanceledException)
        {
            StopWithError("Cancelled", "The rip was cancelled");
            return RipResult.Cancelled;
        }
        catch (Exception ex)
        {
            StopWithError("Unexpected error", ex.Message);
            return RipResult.Failed;
        }
    }

    private void WriteToQueue()
    {
        _logger.Log("QUEUE MODE: Writing encoding job to queue file...");
        _ui.Write("\n[QUEUE MODE] Adding encoding job to queue...", MessageKind.Success);

        var mkvFiles = Directory.GetFiles(_makemkvOutputDir, "*.mkv");

        var queue = new RipQueue(_config.QueueFilePath);
        var count = queue.Append(new QueueEntry
        {
            Title = _options.Title,
            Series = _options.Series,
            Season = _options.Season,
            Disc = _options.Disc,
            OutputDrive = _outputDriveLetter,
            MakeMkvOutputDir = _makemkvOutputDir,
            Bluray = _options.Bluray,
            QueuedAt = DateTime.Now
        });

        _ui.Write("");
        WriteSeparator();
        _ui.Write("QUEUED!", MessageKind.Success);
        WriteSeparator();
        _ui.Write($"\nTitle: {_options.Title}");
        _ui.Write($"MKV files: {mkvFiles.Length}");
        _ui.Write($"Queue file: {queue.QueueFilePath}");
        _ui.Write($"Total jobs in queue: {count}");
        _ui.Write("");
        _ui.Write("Run 'RipDisc -processQueue' to encode all queued jobs sequentially", MessageKind.Warning);
        WriteSeparator();
        _ui.Write("");

        _logger.Log($"QUEUE MODE: Job added to queue ({count} total jobs)");
        _logger.Log($"Queue file: {queue.QueueFilePath}");
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

    /// <summary>Shows what is about to be ripped and asks to go ahead. False means abort.</summary>
    private bool ConfirmStart()
    {
        var driveDescription = _options.DriveIndex >= 0
            ? $"Drive Index {_options.DriveIndex} ({GetDriveHint(_options.DriveIndex)})"
            : $"Drive {_driveLetter}";

        // Validate title doesn't contain metadata that should be separate parameters
        if (!ConfirmTitle())
            return false;

        _ui.Write("");
        WriteSeparator();
        _ui.Write($"Ready to rip: {_options.Title}");

        if (_options.Series)
        {
            if (_options.Season > 0)
            {
                var seasonTag = $"S{_options.Season:D2}";
                _ui.Write($"Type: TV Series - Season {_options.Season} ({seasonTag}), Disc {_options.Disc}");
            }
            else
            {
                _ui.Write($"Type: TV Series - Disc {_options.Disc} (no season folder)");
            }
        }
        else
        {
            var discType = _options.Disc == 1 ? "Main Feature" : "Special Features";
            _ui.Write($"Type: Movie - {discType} (Disc {_options.Disc})");
        }

        _ui.Write($"Using: {driveDescription}", MessageKind.Warning);
        _ui.Write($"Output Drive: {_outputDriveLetter}", MessageKind.Warning);
        WriteSeparator();

        _ui.SetStatus("rip-disc - INPUT");
        if (_ui.Confirm("Start the rip?", defaultAnswer: true))
            return true;

        _ui.Write("Aborted.", MessageKind.Warning);
        _logger.Log("Aborted at the start confirmation");
        return false;
    }

    private bool ConfirmTitle()
    {
        var warnings = TitleValidator.GetMisplacedMetadataWarnings(_options.Title, _options.Series);
        if (warnings.Count == 0)
            return true;

        _ui.Write("");
        WriteSeparator();
        _ui.Write("WARNING: Title may contain misplaced metadata", MessageKind.Error);
        WriteSeparator();
        _ui.Write($"Title: \"{_options.Title}\"", MessageKind.Warning);
        foreach (var w in warnings)
            _ui.Write($"  ! {w}", MessageKind.Warning);

        _ui.Write("");
        _ui.Write("Expected usage:", MessageKind.Header);
        _ui.Write("  RipDisc -title \"Fargo\" -series -season 1 -disc 2");
        _ui.Write("");

        if (_ui.Confirm("Continue with this title?", defaultAnswer: false))
            return true;

        _ui.Write("Aborted. Please re-run with correct parameters.", MessageKind.Warning);
        _logger.Log("Aborted: title may contain misplaced metadata");
        return false;
    }

    private string GetDriveHint(int driveIndex) => _config.GetDriveLabel(driveIndex);

    private void ShowHeader()
    {
        var contentType = _options.Series ? "TV Series" : "Movie";

        _ui.Write("");
        WriteSeparator();
        _ui.Write("DVD/Blu-ray Ripping & Encoding Script", MessageKind.Header);
        WriteSeparator();
        _ui.Write($"Title: {_options.Title}");
        _ui.Write($"Type: {contentType}");

        if (_options.DriveIndex >= 0)
            _ui.Write($"Drive Index: {_options.DriveIndex} ({GetDriveHint(_options.DriveIndex)})");
        else
            _ui.Write($"Drive: {_driveLetter}");

        _ui.Write($"Output Drive: {_outputDriveLetter}");

        if (_options.Series)
        {
            if (_options.Season > 0)
            {
                var seasonTag = $"S{_options.Season:D2}";
                _ui.Write($"Season: {_options.Season} ({seasonTag})");
            }
            else
            {
                _ui.Write("Season: (none - no season folder)");
            }
            _ui.Write($"Disc: {_options.Disc}");
        }
        else
        {
            var discSuffix = _options.Disc > 1 ? " (Special Features)" : "";
            _ui.Write($"Disc: {_options.Disc}{discSuffix}");
        }

        _ui.Write($"MakeMKV Output: {_makemkvOutputDir}");
        _ui.Write($"Final Output: {_finalOutputDir}");
        _ui.Write($"Log file: {_logger.LogFilePath}");
        WriteSeparator();
        _ui.Write("");
    }

    private void EnsureExtrasDirectoryForNonMainDisc()
    {
        if (!_isMainFeatureDisc && !_options.Series)
        {
            var driveCheck = _environment.TestDriveReady(_finalOutputDir);
            if (!driveCheck.Ready)
                throw new ProcessingException("Checking output drive", driveCheck.Message);

            Directory.CreateDirectory(_finalOutputDir);
            Directory.CreateDirectory(_extrasDir);
        }
    }

    /// <summary>Returns false if the user cancelled at the existing-directory prompt.</summary>
    private bool Step1_MakeMKVRip()
    {
        _stepTracker.SetCurrentStep(1);
        _lastWorkingDirectory = _makemkvOutputDir;
        _logger.Log("STEP 1/4: Starting MakeMKV rip...");
        _ui.Write("[STEP 1/4] Starting MakeMKV rip...", MessageKind.Success);

        // Determine disc source
        string discSource;
        if (_options.DriveIndex >= 0)
        {
            discSource = $"disc:{_options.DriveIndex}";
            _ui.Write($"Using drive index: {_options.DriveIndex} (bypasses drive enumeration)", MessageKind.Success);
        }
        else
        {
            discSource = $"dev:{_driveLetter}";
            _ui.Write($"Using drive: {_driveLetter} (may enumerate other drives)", MessageKind.Warning);
            _ui.Write("Tip: Use -DriveIndex to bypass drive enumeration", MessageKind.Detail);
        }

        // Handle existing directory
        _ui.Write($"Creating directory: {_makemkvOutputDir}", MessageKind.Warning);
        if (!HandleExistingMakeMKVDirectory())
            return false;

        // Execute MakeMKV
        var arguments = $"mkv {discSource} all \"{_makemkvOutputDir}\" --minlength=120";
        _ui.Write("\nExecuting MakeMKV command...", MessageKind.Warning);
        _ui.Write($"Command: makemkvcon {arguments}", MessageKind.Detail);
        _logger.Log($"MakeMKV command: makemkvcon {arguments}");

        var (exitCode, output) = _runner.Run(_config.MakeMkvPath, arguments, Step1Label,
            _ui.WriteProcessOutput, MakeMkvTimeout, _cancellationToken);

        // Check for errors
        if (exitCode != 0)
        {
            var driveHint = _options.DriveIndex >= 0 ? GetDriveHint(_options.DriveIndex) : null;
            throw new ProcessingException(Step1Label,
                MakeMkvErrorAnalyzer.AnalyzeExitError(exitCode, output, _options.DriveIndex, _driveLetter, driveHint));
        }

        // Verify files were created
        var rippedFiles = Directory.Exists(_makemkvOutputDir)
            ? Directory.GetFiles(_makemkvOutputDir, "*.mkv")
            : Array.Empty<string>();

        if (rippedFiles.Length == 0)
            throw new ProcessingException(Step1Label, MakeMkvErrorAnalyzer.AnalyzeNoFilesError(output));

        // Show success
        _ui.Write("");
        _ui.Write("MakeMKV rip complete!", MessageKind.Success);
        _ui.Write($"Files ripped: {rippedFiles.Length}");
        _logger.Log($"STEP 1/4: MakeMKV rip complete - {rippedFiles.Length} file(s)");

        foreach (var file in rippedFiles)
        {
            var fileInfo = new FileInfo(file);
            _ui.Write($"  - {fileInfo.Name} ({ToGB(fileInfo.Length)} GB)", MessageKind.Detail);
            _logger.Log($"  Ripped: {fileInfo.Name} ({ToGB(fileInfo.Length)} GB)");
        }

        _stepTracker.CompleteCurrentStep();

        // Eject disc
        _ui.Write($"\nEjecting disc from drive {_driveLetter}...", MessageKind.Warning);
        _environment.EjectDrive(_driveLetter);
        _ui.Write("Disc ejected successfully", MessageKind.Success);
        _logger.Log($"Disc ejected from drive {_driveLetter}");
        return true;
    }

    /// <summary>Returns false if the user cancelled the choice.</summary>
    private bool HandleExistingMakeMKVDirectory()
    {
        if (!Directory.Exists(_makemkvOutputDir))
        {
            Directory.CreateDirectory(_makemkvOutputDir);
            _ui.Write("Directory created successfully", MessageKind.Success);
            return true;
        }

        var existingFiles = Directory.GetFiles(_makemkvOutputDir);
        if (existingFiles.Length == 0)
        {
            _ui.Write("Directory exists (empty)", MessageKind.Detail);
            return true;
        }

        // Directory exists with files
        _ui.Write($"\nWARNING: Directory already exists with {existingFiles.Length} file(s):", MessageKind.Warning);
        _ui.Write($"  {_makemkvOutputDir}");

        foreach (var file in existingFiles)
        {
            var fileInfo = new FileInfo(file);
            _ui.Write($"  - {fileInfo.Name} ({ToGB(fileInfo.Length)} GB)", MessageKind.Detail);
        }

        // Find next available suffix
        int suffix = 1;
        string suffixedDir;
        do
        {
            suffixedDir = $"{_makemkvOutputDir}-{suffix}";
            suffix++;
        } while (Directory.Exists(suffixedDir));

        var choice = _ui.Choose("Choose an option:", new[]
        {
            "Delete existing files and reuse directory",
            $"Use suffixed directory: {suffixedDir}"
        });

        if (choice == 0)
        {
            _ui.Write("Deleting existing files...", MessageKind.Warning);
            foreach (var file in existingFiles)
                File.Delete(file);
            _ui.Write($"Deleted {existingFiles.Length} existing file(s)", MessageKind.Success);
            _logger.Log($"User chose to delete {existingFiles.Length} existing file(s) in {_makemkvOutputDir}");
            return true;
        }

        if (choice == 1)
        {
            _makemkvOutputDir = suffixedDir;

            Directory.CreateDirectory(suffixedDir);
            _ui.Write($"Using suffixed directory: {suffixedDir}", MessageKind.Success);
            _logger.Log($"User chose suffixed directory: {suffixedDir}");
            return true;
        }

        _ui.Write("Aborted - existing files left untouched.", MessageKind.Warning);
        _logger.Log("Aborted at the existing-directory prompt");
        return false;
    }

    private void Step2_HandBrakeEncoding()
    {
        _stepTracker.SetCurrentStep(2);
        _lastWorkingDirectory = _finalOutputDir;
        _logger.Log("STEP 2/4: Starting HandBrake encoding...");
        _ui.Write("\n[STEP 2/4] Starting HandBrake encoding...", MessageKind.Success);

        // Check destination drive
        _ui.Write("Checking destination drive...", MessageKind.Warning);
        var driveCheck = _environment.TestDriveReady(_finalOutputDir);
        if (!driveCheck.Ready)
            throw new ProcessingException(Step2Label, driveCheck.Message);
        _ui.Write($"Destination drive {driveCheck.Drive} is ready", MessageKind.Success);

        // Create output directory
        _ui.Write($"Creating directory: {_finalOutputDir}", MessageKind.Warning);
        if (!Directory.Exists(_finalOutputDir))
        {
            Directory.CreateDirectory(_finalOutputDir);
            _ui.Write("Directory created successfully", MessageKind.Success);
        }
        else
        {
            _ui.Write("Directory already exists", MessageKind.Warning);
        }

        // Encode each MKV file
        var mkvFiles = Directory.GetFiles(_makemkvOutputDir, "*.mkv");
        int fileCount = 0;

        foreach (var mkvFile in mkvFiles)
        {
            fileCount++;
            var mkvInfo = new FileInfo(mkvFile);
            var outputFile = Path.Combine(_finalOutputDir, Path.GetFileNameWithoutExtension(mkvFile) + ".mp4");

            _ui.Write("");
            _ui.Write($"--- Encoding file {fileCount} of {mkvFiles.Length} ---", MessageKind.Header);
            _ui.Write($"Input:  {mkvInfo.Name}");
            _ui.Write($"Output: {Path.GetFileNameWithoutExtension(mkvFile)}.mp4");
            _ui.Write($"Size:   {ToGB(mkvInfo.Length)} GB");
            _logger.Log($"Encoding file {fileCount} of {mkvFiles.Length}: {mkvInfo.Name} ({ToGB(mkvInfo.Length)} GB)");

            _ui.Write("\nExecuting HandBrake...", MessageKind.Warning);

            // Always try with subtitles first (not burned in)
            var baseArgs = $"-i \"{mkvFile}\" -o \"{outputFile}\" " +
                          "--preset \"Fast 1080p30\" " +
                          "--all-audio ";
            var argsWithSubs = baseArgs + "--all-subtitles --subtitle-burned=none --verbose=1";

            var (exitCode, _) = _runner.Run(_config.HandBrakePath, argsWithSubs, Step2Label,
                _ui.WriteProcessOutput, cancellationToken: _cancellationToken);

            // For Bluray: if subtitle encoding fails, retry without subtitles (PGS incompatibility)
            if (exitCode != 0 && _options.Bluray)
            {
                _ui.Write("\nBluray subtitle encoding failed - retrying without subtitles...", MessageKind.Warning);
                _logger.Log($"Bluray subtitle encoding failed for {mkvInfo.Name} - retrying without subtitles");

                // Delete partial output if exists
                if (File.Exists(outputFile))
                    File.Delete(outputFile);

                var argsNoSubs = baseArgs + "--verbose=1";
                (exitCode, _) = _runner.Run(_config.HandBrakePath, argsNoSubs, Step2Label,
                    _ui.WriteProcessOutput, cancellationToken: _cancellationToken);
            }

            if (exitCode != 0)
                throw new ProcessingException(Step2Label,
                    $"HandBrake exited with code {exitCode} while encoding {mkvInfo.Name}");

            if (!File.Exists(outputFile))
                throw new ProcessingException(Step2Label,
                    $"Output file not created for {mkvInfo.Name}");

            var outputInfo = new FileInfo(outputFile);
            _ui.Write("");
            _ui.Write($"Encoding complete: {mkvInfo.Name}", MessageKind.Success);
            _ui.Write($"Output size: {ToGB(outputInfo.Length)} GB");
            _logger.Log($"Encoded: {mkvInfo.Name} -> {Path.GetFileNameWithoutExtension(mkvFile)}.mp4 ({ToGB(outputInfo.Length)} GB)");
        }

        _stepTracker.CompleteCurrentStep();
        _logger.Log($"STEP 2/4: HandBrake encoding complete - {fileCount} file(s) encoded");

        // Wait for file handles to be released
        _ui.Write("\nWaiting for file handles to be released...", MessageKind.Warning);
        _environment.WaitForFileHandles();
        _ui.Write("File handle wait complete", MessageKind.Success);

        // Delete MakeMKV temporary directory
        _ui.Write("\nChecking for successful encodes...", MessageKind.Warning);
        var encodedFiles = Directory.GetFiles(_finalOutputDir, "*.mp4");
        if (encodedFiles.Length > 0)
        {
            _ui.Write($"Found {encodedFiles.Length} encoded file(s)", MessageKind.Success);
            _ui.Write($"Removing temporary MakeMKV directory: {_makemkvOutputDir}", MessageKind.Warning);
            Directory.Delete(_makemkvOutputDir, true);
            _ui.Write("Temporary files removed successfully", MessageKind.Success);
            _logger.Log($"Temporary MKV directory removed: {_makemkvOutputDir}");
        }
        else
        {
            _ui.Write("WARNING: No encoded files found. Keeping MakeMKV directory.", MessageKind.Error);
            _logger.Log("WARNING: No encoded files found - keeping MakeMKV directory");
        }
    }

    private void Step3_OrganizeFiles()
    {
        _stepTracker.SetCurrentStep(3);
        _lastWorkingDirectory = _finalOutputDir;
        _logger.Log("STEP 3/4: Organizing files...");
        _ui.Write("\n[STEP 3/4] Organizing files...", MessageKind.Success);

        // Every operation below uses full paths. The console app used to change the
        // process-wide current directory here; a GUI runs this on a background thread, so
        // that is gone.
        _ui.Write($"Working in: {_finalOutputDir}", MessageKind.Warning);

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
        _ui.Write("\nDeleting image files...", MessageKind.Warning);
        var imageExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp" };
        var imageFiles = Directory.GetFiles(_finalOutputDir)
            .Where(f => imageExtensions.Contains(Path.GetExtension(f).ToLower()))
            .ToArray();

        if (imageFiles.Length > 0)
        {
            _ui.Write($"Image files to delete: {imageFiles.Length}");
            foreach (var file in imageFiles)
            {
                _ui.Write($"  - {Path.GetFileName(file)}", MessageKind.Detail);
                File.Delete(file);
            }
            _ui.Write("Image files deleted", MessageKind.Success);
        }
        else
        {
            _ui.Write("No image files found", MessageKind.Detail);
        }
    }

    private void PrefixSeriesFiles()
    {
        _ui.Write("\nPrefixing files with title...", MessageKind.Warning);
        var filesToPrefix = Directory.GetFiles(_finalOutputDir)
            .Where(f => !Path.GetFileName(f).StartsWith($"{_options.Title}-"))
            .ToArray();

        if (filesToPrefix.Length > 0)
        {
            _ui.Write($"Files to prefix: {filesToPrefix.Length}");
            foreach (var file in filesToPrefix)
            {
                var fileName = Path.GetFileName(file);
                _ui.Write($"  - {fileName}", MessageKind.Detail);
                var newPath = Path.Combine(_finalOutputDir, $"{_options.Title}-{fileName}");
                File.Move(file, newPath);
            }
            _ui.Write("Prefixing complete", MessageKind.Success);
            _logger.Log($"Prefixed {filesToPrefix.Length} file(s) with title");
        }
        else
        {
            _ui.Write("No files need prefixing", MessageKind.Detail);
        }
    }

    private void PrefixMovieFiles()
    {
        var dirName = new DirectoryInfo(_finalOutputDir).Name;
        var filesToPrefix = Directory.GetFiles(_finalOutputDir)
            .Where(f => !Path.GetFileName(f).StartsWith($"{dirName}-"))
            .ToArray();

        if (_isMainFeatureDisc)
        {
            _ui.Write("\nPrefixing files with directory name...", MessageKind.Warning);

            if (filesToPrefix.Length > 0)
            {
                _ui.Write($"Files to prefix: {filesToPrefix.Length}");
                foreach (var file in filesToPrefix)
                {
                    var fileName = Path.GetFileName(file);
                    _ui.Write($"  - {fileName}", MessageKind.Detail);
                    var newPath = Path.Combine(_finalOutputDir, $"{dirName}-{fileName}");
                    File.Move(file, newPath);
                }
                _ui.Write("Prefixing complete", MessageKind.Success);
                _logger.Log($"Prefixed {filesToPrefix.Length} file(s) with directory name");
            }
            else
            {
                _ui.Write("No files need prefixing", MessageKind.Detail);
            }
        }
        else
        {
            _ui.Write("\nPrefixing special features files...", MessageKind.Warning);

            if (filesToPrefix.Length > 0)
            {
                _ui.Write($"Files to prefix: {filesToPrefix.Length}");
                foreach (var file in filesToPrefix)
                {
                    var fileName = Path.GetFileName(file);
                    var newName = $"{dirName}-Special Features-{fileName}";
                    _ui.Write($"  - {fileName} -> {newName}", MessageKind.Detail);
                    var newPath = Path.Combine(_finalOutputDir, newName);
                    File.Move(file, newPath);
                }
                _ui.Write("Special features prefixing complete", MessageKind.Success);
                _logger.Log($"Prefixed {filesToPrefix.Length} special features file(s)");
            }
            else
            {
                _ui.Write("No files need prefixing", MessageKind.Detail);
            }
        }
    }

    private void RenameFeatureFile()
    {
        _ui.Write("\nChecking for Feature file...", MessageKind.Warning);
        var featureFile = Directory.GetFiles(_finalOutputDir)
            .FirstOrDefault(f => Path.GetFileName(f).Contains("-Feature."));

        if (featureFile != null)
        {
            _ui.Write($"Feature file already exists: {Path.GetFileName(featureFile)}", MessageKind.Detail);
            return;
        }

        var largestFile = Directory.GetFiles(_finalOutputDir)
            .Select(f => new FileInfo(f))
            .OrderByDescending(f => f.Length)
            .FirstOrDefault();

        if (largestFile == null)
            return;

        var dirName = new DirectoryInfo(_finalOutputDir).Name;
        var newName = $"{dirName}-Feature{largestFile.Extension}";
        var newPath = Path.Combine(_finalOutputDir, newName);

        _ui.Write($"Largest file: {largestFile.Name} ({ToGB(largestFile.Length)} GB)");
        _ui.Write($"Renaming to: {newName}", MessageKind.Warning);
        File.Move(largestFile.FullName, newPath);
        _ui.Write("Feature file renamed successfully", MessageKind.Success);
        _logger.Log($"Feature file: {largestFile.Name} -> {newName} ({ToGB(largestFile.Length)} GB)");
    }

    private static readonly string[] VideoExtensions = { ".mp4", ".avi", ".mkv", ".mov", ".wmv" };

    private void MoveNonFeatureToExtras()
    {
        _ui.Write("\nChecking for non-feature videos...", MessageKind.Warning);
        var nonFeatureVideos = Directory.GetFiles(_finalOutputDir)
            .Where(f => VideoExtensions.Contains(Path.GetExtension(f).ToLower()) &&
                       !Path.GetFileName(f).Contains("Feature"))
            .ToArray();

        if (nonFeatureVideos.Length == 0)
        {
            _ui.Write("No non-feature videos found", MessageKind.Detail);
            return;
        }

        _ui.Write($"Non-feature videos found: {nonFeatureVideos.Length}");
        foreach (var file in nonFeatureVideos)
            _ui.Write($"  - {Path.GetFileName(file)}", MessageKind.Detail);

        EnsureExtrasDirectory();

        _ui.Write("Moving files to extras...", MessageKind.Warning);
        foreach (var file in nonFeatureVideos)
        {
            var destPath = Path.Combine(_extrasDir, Path.GetFileName(file));
            File.Move(file, destPath);
        }
        _ui.Write("Files moved to extras", MessageKind.Success);
        _logger.Log($"Moved {nonFeatureVideos.Length} non-feature file(s) to extras");
    }

    private void MoveSpecialFeaturesToExtras()
    {
        _ui.Write("\nMoving special features to extras folder...", MessageKind.Warning);
        EnsureExtrasDirectory();

        var videoFiles = Directory.GetFiles(_finalOutputDir)
            .Where(f => VideoExtensions.Contains(Path.GetExtension(f).ToLower()) &&
                       !Path.GetFileName(f).Contains("-Feature."))
            .ToArray();

        if (videoFiles.Length == 0)
        {
            _ui.Write("No video files to move", MessageKind.Detail);
            return;
        }

        _ui.Write($"Videos to move: {videoFiles.Length}");
        foreach (var video in videoFiles)
        {
            var fileName = Path.GetFileName(video);
            var uniquePath = FileHelper.GetUniqueFilePath(_extrasDir, fileName);
            var newName = Path.GetFileName(uniquePath);

            if (newName != fileName)
                _ui.Write($"  - {fileName} -> {newName} (renamed to avoid clash)", MessageKind.Warning);
            else
                _ui.Write($"  - {fileName}", MessageKind.Detail);

            File.Move(video, uniquePath);
        }
        _ui.Write("Files moved to extras", MessageKind.Success);
        _logger.Log($"Moved {videoFiles.Length} special features file(s) to extras");
    }

    private void EnsureExtrasDirectory()
    {
        if (!Directory.Exists(_extrasDir))
        {
            _ui.Write("Creating extras directory...", MessageKind.Warning);
            Directory.CreateDirectory(_extrasDir);
            _ui.Write("Extras directory created", MessageKind.Success);
        }
        else
        {
            _ui.Write("Extras directory already exists", MessageKind.Detail);
        }
    }

    private void Step4_OpenDirectory()
    {
        _stepTracker.SetCurrentStep(4);
        _logger.Log("STEP 4/4: Opening directory...");
        _ui.Write("\n[STEP 4/4] Opening film directory...", MessageKind.Success);
        _ui.Write($"Opening: {_finalOutputDir}", MessageKind.Warning);
        _environment.OpenDirectory(_finalOutputDir);
        _stepTracker.CompleteCurrentStep();
    }

    private void ShowCompletionSummary()
    {
        _ui.Write("");
        WriteSeparator();
        _ui.Write("COMPLETE!", MessageKind.Success);
        WriteSeparator();

        _ui.Write($"\nProcessed: {GetTitleSummary()}");
        _ui.Write($"Final location: {_finalOutputDir}");

        _stepTracker.WriteSummary(_ui);

        // File summary
        _ui.Write("");
        _ui.Write("--- FILE SUMMARY ---", MessageKind.Header);
        var finalFiles = Directory.GetFiles(_finalOutputDir, "*", SearchOption.AllDirectories);
        var totalSize = finalFiles.Sum(f => new FileInfo(f).Length);
        var totalSizeGB = ToGB(totalSize);

        _ui.Write($"  Total files: {finalFiles.Length}");
        _ui.Write($"  Total size: {totalSizeGB} GB");
        _ui.Write($"  Log file: {_logger.LogFilePath}");
        WriteSeparator();
        _ui.Write("");

        _logger.Log("========== RIP SESSION COMPLETE ==========");
        _logger.Log($"Final location: {_finalOutputDir}");
        _logger.Log($"Total files: {finalFiles.Length}");
        _logger.Log($"Total size: {totalSizeGB} GB");
        foreach (var file in finalFiles)
        {
            var fileInfo = new FileInfo(file);
            _logger.Log($"  {fileInfo.Name} ({ToGB(fileInfo.Length)} GB)");
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
        _ui.SetStatus($"{_windowTitle} - ERROR");

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
        _ui.Write("");
        WriteSeparator();
        _ui.Write("FAILED!", MessageKind.Error);
        WriteSeparator();

        _ui.Write($"\nProcessing: {GetTitleSummary()}");
        _ui.Write("");
        _ui.Write($"Error at: {step}", MessageKind.Error);
        _ui.Write($"Message: {message}", MessageKind.Error);

        _stepTracker.WriteSummary(_ui, showRemaining: true);

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
            _ui.Write("");
            _ui.Write("--- OPENING DIRECTORY ---", MessageKind.Header);
            _ui.Write($"Opening: {directoryToOpen}", MessageKind.Warning);
            _ui.Write("(This is where leftover/partial files may be located)", MessageKind.Detail);
            _environment.OpenDirectory(directoryToOpen);
        }

        _ui.Write($"\nLog file: {_logger.LogFilePath}", MessageKind.Warning);
        _ui.Write("");
        WriteSeparator();
        _ui.Write("Please complete the remaining steps manually", MessageKind.Error);
        WriteSeparator();
        _ui.Write("");
    }

    private void ShowManualSteps(List<ProcessingStep> remainingSteps)
    {
        _ui.Write("");
        _ui.Write("--- MANUAL STEPS NEEDED ---", MessageKind.Header);

        foreach (var step in remainingSteps)
        {
            switch (step.Number)
            {
                case 1:
                    _ui.Write("  - Re-run MakeMKV to rip the disc", MessageKind.Warning);
                    break;
                case 2:
                    _ui.Write("  - Encode MKV files with HandBrake", MessageKind.Warning);
                    if (Directory.Exists(_makemkvOutputDir))
                        _ui.Write($"    MKV files location: {_makemkvOutputDir}", MessageKind.Detail);
                    break;
                case 3:
                    _ui.Write("  - Rename files to proper format", MessageKind.Warning);
                    if (_options.Series)
                    {
                        _ui.Write($"    Format: {_options.Title}-originalname.mp4", MessageKind.Detail);
                    }
                    else if (_options.Disc == 1)
                    {
                        _ui.Write($"    Format: {_options.Title}-Feature.mp4 (largest file)", MessageKind.Detail);
                        _ui.Write($"    Move extras to: {_extrasDir}", MessageKind.Detail);
                    }
                    else
                    {
                        _ui.Write($"    Format: {_options.Title}-Special Features-originalname.mp4", MessageKind.Detail);
                        _ui.Write($"    Move all files to: {_extrasDir}", MessageKind.Detail);
                    }
                    break;
                case 4:
                    _ui.Write("  - Open output directory to verify files", MessageKind.Warning);
                    break;
            }
        }
    }

    private void WriteSeparator() => _ui.Write("========================================", MessageKind.Header);

    private static double ToGB(long bytes) => Math.Round(bytes / (1024.0 * 1024.0 * 1024.0), 2);
}

public class ProcessingException : Exception
{
    public string Step { get; }

    public ProcessingException(string step, string message) : base(message)
    {
        Step = step;
    }
}

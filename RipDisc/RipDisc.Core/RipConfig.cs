using System.Text.Json;
using System.Text.Json.Serialization;

namespace RipDisc.Core;

/// <summary>
/// The C# side of ripdisc-config.json - the same file Load-Config.ps1 reads, with the same
/// defaults and the same tool-path auto-detection fallbacks.
/// </summary>
public class RipConfig
{
    public const string FileName = "ripdisc-config.json";

    private static readonly string[] MakeMkvCandidates =
    {
        @"C:\Program Files (x86)\MakeMKV\makemkvcon64.exe",
        @"C:\Program Files\MakeMKV\makemkvcon64.exe",
        @"C:\Program Files (x86)\MakeMKV\makemkvcon.exe",
        @"C:\Program Files\MakeMKV\makemkvcon.exe"
    };

    private static readonly string[] HandBrakeCandidates =
    {
        @"C:\ProgramData\chocolatey\bin\HandBrakeCLI.exe",
        @"C:\Program Files\HandBrake\HandBrakeCLI.exe",
        @"C:\Program Files (x86)\HandBrake\HandBrakeCLI.exe"
    };

    public string MakeMkvPath { get; set; } = string.Empty;
    public string HandBrakePath { get; set; } = string.Empty;
    public string TempRoot { get; set; } = @"C:\Video";
    public string DefaultInputDrive { get; set; } = "D:";
    public string DefaultOutputDrive { get; set; } = "E:";
    public Dictionary<string, string> DriveLabels { get; set; } = new();
    public string TmdbApiKey { get; set; } = string.Empty;

    /// <summary>The config file that was loaded, or null if defaults were used.</summary>
    [JsonIgnore]
    public string? SourcePath { get; private set; }

    [JsonIgnore]
    public string LogDirectory => Path.Combine(TempRoot, "logs");

    [JsonIgnore]
    public string QueueFilePath => Path.Combine(TempRoot, "handbrake-queue.json");

    public string GetDriveLabel(int driveIndex) =>
        DriveLabels.TryGetValue(driveIndex.ToString(), out var label) && !string.IsNullOrWhiteSpace(label)
            ? label
            : "unknown drive";

    /// <summary>
    /// Loads the config from <paramref name="explicitPath"/>, or else the first
    /// ripdisc-config.json found walking up from the executable's folder (so a dev build
    /// under RipDisc\...\bin finds the repo-root config). Missing file means defaults.
    /// </summary>
    public static RipConfig Load(string? explicitPath = null)
    {
        var path = explicitPath ?? FindConfigFile(AppContext.BaseDirectory);

        RipConfig config;
        if (path != null && File.Exists(path))
        {
            config = Parse(File.ReadAllText(path));
            config.SourcePath = path;
        }
        else
        {
            config = new RipConfig();
        }

        config.ResolveToolPaths();
        return config;
    }

    public static string? FindConfigFile(string startDirectory)
    {
        var dir = new DirectoryInfo(startDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(dir.FullName, FileName);
            if (File.Exists(candidate))
                return candidate;
            dir = dir.Parent;
        }
        return null;
    }

    internal static RipConfig Parse(string json)
    {
        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true
        };
        var config = JsonSerializer.Deserialize<RipConfig>(json, options) ?? new RipConfig();

        // Blank values in the file mean "use the default", same as Load-Config.ps1
        var defaults = new RipConfig();
        if (string.IsNullOrWhiteSpace(config.TempRoot)) config.TempRoot = defaults.TempRoot;
        if (string.IsNullOrWhiteSpace(config.DefaultInputDrive)) config.DefaultInputDrive = defaults.DefaultInputDrive;
        if (string.IsNullOrWhiteSpace(config.DefaultOutputDrive)) config.DefaultOutputDrive = defaults.DefaultOutputDrive;
        config.DriveLabels ??= new Dictionary<string, string>();
        config.MakeMkvPath ??= string.Empty;
        config.HandBrakePath ??= string.Empty;
        config.TmdbApiKey ??= string.Empty;
        return config;
    }

    private void ResolveToolPaths()
    {
        if (string.IsNullOrWhiteSpace(MakeMkvPath) || !File.Exists(MakeMkvPath))
            MakeMkvPath = FindTool(MakeMkvCandidates) ?? MakeMkvPath;

        if (string.IsNullOrWhiteSpace(HandBrakePath) || !File.Exists(HandBrakePath))
            HandBrakePath = FindTool(HandBrakeCandidates) ?? HandBrakePath;

        if (DriveLabels.Count == 0)
        {
            DriveLabels["0"] = "Internal drive";
            DriveLabels["1"] = "External drive";
        }

        if (string.IsNullOrWhiteSpace(TmdbApiKey))
            TmdbApiKey = Environment.GetEnvironmentVariable("TMDB_API_KEY") ?? string.Empty;
    }

    private static string? FindTool(string[] candidates)
    {
        var pathDirs = (Environment.GetEnvironmentVariable("PATH") ?? string.Empty)
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);

        foreach (var exeName in candidates.Select(Path.GetFileName).Distinct())
        {
            foreach (var dir in pathDirs)
            {
                try
                {
                    var candidate = Path.Combine(dir.Trim(), exeName!);
                    if (File.Exists(candidate))
                        return candidate;
                }
                catch (ArgumentException)
                {
                    // Malformed PATH entry - skip it
                }
            }
        }

        return candidates.FirstOrDefault(File.Exists);
    }
}

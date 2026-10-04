namespace RipDisc.Core;

/// <summary>Where a rip's files go, worked out once from the options and config.</summary>
public class RipPaths
{
    public required string DriveLetter { get; init; }
    public required string OutputDriveLetter { get; init; }
    public required string MakeMkvOutputDir { get; init; }
    public required string FinalOutputDir { get; init; }
    public string ExtrasDir => Path.Combine(FinalOutputDir, "extras");

    public static RipPaths Create(RipOptions options, RipConfig config)
    {
        var driveLetter = Core.DriveLetter.Normalize(
            string.IsNullOrWhiteSpace(options.Drive) ? config.DefaultInputDrive : options.Drive);
        var outputDriveLetter = Core.DriveLetter.Normalize(
            string.IsNullOrWhiteSpace(options.OutputDrive) ? config.DefaultOutputDrive : options.OutputDrive);

        var makemkvOutputDir = options.Disc > 1
            ? Path.Combine(config.TempRoot, options.Title, $"Disc{options.Disc}")
            : Path.Combine(config.TempRoot, options.Title);

        string finalOutputDir;
        if (options.Series)
        {
            var seriesBaseDir = $@"{outputDriveLetter}\Series\{options.Title}";
            finalOutputDir = options.Season > 0
                ? Path.Combine(seriesBaseDir, $"Season {options.Season}")
                : seriesBaseDir;
        }
        else
        {
            finalOutputDir = $@"{outputDriveLetter}\DVDs\{options.Title}";
        }

        return new RipPaths
        {
            DriveLetter = driveLetter,
            OutputDriveLetter = outputDriveLetter,
            MakeMkvOutputDir = makemkvOutputDir,
            FinalOutputDir = finalOutputDir
        };
    }
}

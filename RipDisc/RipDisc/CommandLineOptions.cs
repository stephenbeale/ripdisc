namespace RipDisc;

public class CommandLineOptions
{
    public string Title { get; set; } = string.Empty;
    public bool Series { get; set; }
    public int Season { get; set; }
    public int Disc { get; set; } = 1;
    public string Drive { get; set; } = "D:";
    public int DriveIndex { get; set; } = -1;
    public string OutputDrive { get; set; } = "E:";
}

public static class CommandLineParser
{
    public static CommandLineOptions Parse(string[] args)
    {
        var options = new CommandLineOptions();

        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i].ToLower();

            switch (arg)
            {
                case "-title":
                    if (i + 1 >= args.Length)
                        throw new ArgumentException("Missing value for -title");
                    options.Title = args[++i];
                    break;

                case "-series":
                    options.Series = true;
                    break;

                case "-season":
                    if (i + 1 >= args.Length)
                        throw new ArgumentException("Missing value for -season");
                    if (!int.TryParse(args[++i], out int season))
                        throw new ArgumentException("Invalid value for -season");
                    options.Season = season;
                    break;

                case "-disc":
                    if (i + 1 >= args.Length)
                        throw new ArgumentException("Missing value for -disc");
                    if (!int.TryParse(args[++i], out int disc))
                        throw new ArgumentException("Invalid value for -disc");
                    options.Disc = disc;
                    break;

                case "-drive":
                    if (i + 1 >= args.Length)
                        throw new ArgumentException("Missing value for -drive");
                    options.Drive = args[++i];
                    break;

                case "-driveindex":
                    if (i + 1 >= args.Length)
                        throw new ArgumentException("Missing value for -driveIndex");
                    if (!int.TryParse(args[++i], out int driveIndex))
                        throw new ArgumentException("Invalid value for -driveIndex");
                    options.DriveIndex = driveIndex;
                    break;

                case "-outputdrive":
                    if (i + 1 >= args.Length)
                        throw new ArgumentException("Missing value for -outputDrive");
                    options.OutputDrive = args[++i];
                    break;

                default:
                    throw new ArgumentException($"Unknown argument: {args[i]}");
            }
        }

        if (string.IsNullOrWhiteSpace(options.Title))
            throw new ArgumentException("Title is required");

        // Normalize drive letters
        if (!options.Drive.EndsWith(":"))
            options.Drive += ":";
        if (!options.OutputDrive.EndsWith(":"))
            options.OutputDrive += ":";

        return options;
    }
}

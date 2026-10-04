using RipDisc.Core;

namespace RipDisc;

public class CommandLineOptions : RipOptions
{
    public bool ProcessQueue { get; set; }
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
                    options.Title = RequireValue(args, ref i, "-title");
                    break;

                case "-series":
                    options.Series = true;
                    break;

                case "-season":
                    options.Season = RequireInt(args, ref i, "-season");
                    break;

                case "-disc":
                    options.Disc = RequireInt(args, ref i, "-disc");
                    break;

                case "-drive":
                    options.Drive = DriveLetter.Normalize(RequireValue(args, ref i, "-drive"));
                    break;

                case "-driveindex":
                    options.DriveIndex = RequireInt(args, ref i, "-driveIndex");
                    break;

                case "-outputdrive":
                    options.OutputDrive = DriveLetter.Normalize(RequireValue(args, ref i, "-outputDrive"));
                    break;

                case "-queue":
                    options.Queue = true;
                    break;

                case "-processqueue":
                    options.ProcessQueue = true;
                    break;

                case "-bluray":
                    options.Bluray = true;
                    break;

                default:
                    throw new ArgumentException($"Unknown argument: {args[i]}");
            }
        }

        if (options.ProcessQueue && options.Queue)
            throw new ArgumentException("-queue and -processQueue are mutually exclusive");

        if (options.ProcessQueue)
            return options;

        if (string.IsNullOrWhiteSpace(options.Title))
            throw new ArgumentException("Title is required");

        return options;
    }

    private static string RequireValue(string[] args, ref int i, string name)
    {
        if (i + 1 >= args.Length)
            throw new ArgumentException($"Missing value for {name}");
        return args[++i];
    }

    private static int RequireInt(string[] args, ref int i, string name)
    {
        if (!int.TryParse(RequireValue(args, ref i, name), out int value))
            throw new ArgumentException($"Invalid value for {name}");
        return value;
    }
}

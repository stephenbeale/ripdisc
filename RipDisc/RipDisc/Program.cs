namespace RipDisc;

class Program
{
    static int Main(string[] args)
    {
        try
        {
            var options = CommandLineParser.Parse(args);

            if (options.ProcessQueue)
                return RipDiscApplication.ProcessAllQueued();

            var app = new RipDiscApplication(options);
            return app.Run();
        }
        catch (ArgumentException ex)
        {
            ConsoleHelper.WriteError(ex.Message);
            ShowUsage();
            return 1;
        }
        catch (Exception ex)
        {
            ConsoleHelper.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    static void ShowUsage()
    {
        Console.WriteLine();
        Console.WriteLine("Usage: RipDisc -title <title> [options]");
        Console.WriteLine("       RipDisc -processQueue");
        Console.WriteLine();
        Console.WriteLine("Required:");
        Console.WriteLine("  -title <string>        Title of the movie or series");
        Console.WriteLine();
        Console.WriteLine("Optional:");
        Console.WriteLine("  -series                Flag for TV series (no value needed)");
        Console.WriteLine("  -season <int>          Season number (default: 0)");
        Console.WriteLine("  -disc <int>            Disc number (default: 1)");
        Console.WriteLine("  -drive <string>        Drive letter (default: D:)");
        Console.WriteLine("  -driveIndex <int>      Drive index for MakeMKV (default: -1)");
        Console.WriteLine("  -outputDrive <string>  Output drive letter (default: E:)");
        Console.WriteLine("  -queue                 Queue encoding instead of running inline");
        Console.WriteLine("  -bluray                Skip subtitles (Blu-ray PGS subs don't work in MP4)");
        Console.WriteLine();
        Console.WriteLine("Queue Mode:");
        Console.WriteLine("  -processQueue          Process all queued encoding jobs sequentially");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  RipDisc -title \"The Matrix\"");
        Console.WriteLine("  RipDisc -title \"Breaking Bad\" -series -season 1 -disc 1");
        Console.WriteLine("  RipDisc -title \"The Matrix\" -disc 2 -driveIndex 1");
        Console.WriteLine();
        Console.WriteLine("  RipDisc -title \"The Matrix\" -queue                  # Rip + queue encode");
        Console.WriteLine("  RipDisc -title \"The Matrix\" -disc 2 -queue          # Rip disc 2 + queue");
        Console.WriteLine("  RipDisc -processQueue                                # Encode all queued");
    }
}

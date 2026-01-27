namespace RipDisc;

public class Logger
{
    private readonly string _logFilePath;

    public Logger(string title, int disc)
    {
        var logDir = @"C:\Video\logs";
        Directory.CreateDirectory(logDir);

        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        _logFilePath = Path.Combine(logDir, $"{title}_disc{disc}_{timestamp}.log");
    }

    public string LogFilePath => _logFilePath;

    public void Log(string message)
    {
        try
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var entry = $"[{timestamp}] {message}";
            File.AppendAllText(_logFilePath, entry + Environment.NewLine);
        }
        catch
        {
            // Silently ignore logging errors
        }
    }
}

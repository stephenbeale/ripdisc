namespace RipDisc.Core;

public class Logger
{
    private readonly string _logFilePath;

    public Logger(string logDirectory, string title, int disc)
    {
        Directory.CreateDirectory(logDirectory);

        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        _logFilePath = Path.Combine(logDirectory, $"{title}_disc{disc}_{timestamp}.log");
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

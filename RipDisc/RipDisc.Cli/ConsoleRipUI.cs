using RipDisc.Core;

namespace RipDisc;

/// <summary>The console front end for the pipeline.</summary>
public class ConsoleRipUI : IRipUI
{
    private readonly object _consoleLock = new();

    public void Write(string message, MessageKind kind = MessageKind.Info)
    {
        lock (_consoleLock)
        {
            switch (kind)
            {
                case MessageKind.Header: ConsoleHelper.WriteHeader(message); break;
                case MessageKind.Success: ConsoleHelper.WriteSuccess(message); break;
                case MessageKind.Error: ConsoleHelper.WriteError(message); break;
                case MessageKind.Warning: ConsoleHelper.WriteWarning(message); break;
                case MessageKind.Detail: ConsoleHelper.WriteGray(message); break;
                default: ConsoleHelper.WriteInfo(message); break;
            }
        }
    }

    public void WriteProcessOutput(string line)
    {
        lock (_consoleLock)
            Console.WriteLine(line);
    }

    public void SetStatus(string status) => ConsoleHelper.SetWindowTitle(status);

    public bool Confirm(string prompt, bool defaultAnswer)
    {
        var hint = defaultAnswer ? "(Y/n)" : "(y/N)";
        while (true)
        {
            var response = ReadLine($"{prompt} {hint}: ");

            // Closed stdin is never consent - a piped blank line once started a real encode
            if (response == null)
                return false;

            response = response.Trim();
            if (response.Length == 0)
                return defaultAnswer;
            if (response.Equals("y", StringComparison.OrdinalIgnoreCase) || response.Equals("yes", StringComparison.OrdinalIgnoreCase))
                return true;
            if (response.Equals("n", StringComparison.OrdinalIgnoreCase) || response.Equals("no", StringComparison.OrdinalIgnoreCase))
                return false;

            Write("Please answer y or n.", MessageKind.Error);
        }
    }

    public int Choose(string prompt, IReadOnlyList<string> options)
    {
        Write("\n" + prompt, MessageKind.Header);
        for (int i = 0; i < options.Count; i++)
            Write($"  [{i + 1}] {options[i]}", MessageKind.Warning);

        while (true)
        {
            var response = ReadLine($"Enter 1-{options.Count}: ");
            if (response == null)
                return -1;

            if (int.TryParse(response.Trim(), out var number) && number >= 1 && number <= options.Count)
                return number - 1;

            Write($"Invalid choice. Please enter a number from 1 to {options.Count}.", MessageKind.Error);
        }
    }

    private string? ReadLine(string prompt)
    {
        lock (_consoleLock)
            Console.Write(prompt);
        return Console.ReadLine();
    }
}

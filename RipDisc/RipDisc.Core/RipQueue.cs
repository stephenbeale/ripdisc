using System.Text.Json;

namespace RipDisc.Core;

public class QueueEntry
{
    public string Title { get; set; } = string.Empty;
    public bool Series { get; set; }
    public int Season { get; set; }
    public int Disc { get; set; } = 1;
    public string OutputDrive { get; set; } = "E:";
    public string MakeMkvOutputDir { get; set; } = string.Empty;
    public bool Bluray { get; set; }
    public DateTime QueuedAt { get; set; }
}

/// <summary>The shared handbrake-queue.json file that -queue rips append to.</summary>
public class RipQueue
{
    private static readonly JsonSerializerOptions WriteOptions = new() { WriteIndented = true };

    public RipQueue(string queueFilePath)
    {
        QueueFilePath = queueFilePath;
    }

    public string QueueFilePath { get; }

    public bool Exists => File.Exists(QueueFilePath);

    public List<QueueEntry> Read()
    {
        if (!File.Exists(QueueFilePath))
            return new List<QueueEntry>();
        return JsonSerializer.Deserialize<List<QueueEntry>>(File.ReadAllText(QueueFilePath)) ?? new List<QueueEntry>();
    }

    /// <summary>Appends an entry and returns the queue length afterwards.</summary>
    public int Append(QueueEntry entry)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(QueueFilePath)!);

        // Lock file protects concurrent writes from parallel rip sessions
        var lockPath = QueueFilePath + ".lock";
        int count;
        using (new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
        {
            var queue = Read();
            queue.Add(entry);
            Write(queue);
            count = queue.Count;
        }

        try { File.Delete(lockPath); } catch { }
        return count;
    }

    /// <summary>Writes the queue back, deleting the file when it is empty.</summary>
    public void Write(List<QueueEntry> queue)
    {
        if (queue.Count > 0)
            File.WriteAllText(QueueFilePath, JsonSerializer.Serialize(queue, WriteOptions));
        else if (File.Exists(QueueFilePath))
            File.Delete(QueueFilePath);
    }
}

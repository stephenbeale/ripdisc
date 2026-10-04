namespace RipDisc.Core;

/// <summary>Encodes every job in the shared queue file, one at a time.</summary>
public class QueueProcessor
{
    private readonly RipConfig _config;
    private readonly IRipUI _ui;
    private readonly Func<RipOptions, QueueEntry, RipPipeline> _createPipeline;

    public QueueProcessor(RipConfig config, IRipUI ui)
        : this(config, ui, (options, _) => new RipPipeline(options, config, ui))
    {
    }

    internal QueueProcessor(RipConfig config, IRipUI ui, Func<RipOptions, QueueEntry, RipPipeline> createPipeline)
    {
        _config = config;
        _ui = ui;
        _createPipeline = createPipeline;
    }

    public RipResult ProcessAll(CancellationToken cancellationToken = default)
    {
        _ui.SetStatus("RipDisc - Processing Queue");
        var store = new RipQueue(_config.QueueFilePath);
        int completedCount = 0;

        if (!store.Exists)
        {
            _ui.Write("No queue file found.", MessageKind.Error);
            _ui.Write($"Expected: {store.QueueFilePath}", MessageKind.Detail);
            _ui.Write("Use -queue flag when ripping to add jobs to the queue.", MessageKind.Detail);
            return RipResult.Failed;
        }

        var queue = store.Read();
        if (queue.Count == 0)
        {
            _ui.Write("Queue is empty - nothing to process.", MessageKind.Warning);
            return RipResult.Success;
        }

        _ui.Write("");
        WriteSeparator();
        _ui.Write($"PROCESSING QUEUE: {queue.Count} job(s)", MessageKind.Success);
        WriteSeparator();

        var done = new List<QueueEntry>();

        while (queue.Count > 0)
        {
            var entry = queue[0];
            completedCount++;

            _ui.Write("");
            _ui.Write($"--- Job {completedCount} of {queue.Count + completedCount - 1}: {entry.Title} (Disc {entry.Disc}) ---", MessageKind.Header);

            var options = new RipOptions
            {
                Title = entry.Title,
                Series = entry.Series,
                Season = entry.Season,
                Disc = entry.Disc,
                OutputDrive = entry.OutputDrive,
                Bluray = entry.Bluray
            };

            var pipeline = _createPipeline(options, entry);

            // Override MakeMKV output dir with the actual path recorded at queue time
            if (!string.IsNullOrEmpty(entry.MakeMkvOutputDir))
                pipeline.SetMakeMkvOutputDir(entry.MakeMkvOutputDir);

            var result = pipeline.RunFromQueue(cancellationToken);

            if (result != RipResult.Success)
            {
                _ui.Write($"\nJob failed: {entry.Title}", MessageKind.Error);
                _ui.Write("Stopping queue processing. Remaining jobs preserved in queue file.", MessageKind.Warning);
                store.Write(queue);
                return result;
            }

            // Remove completed job and re-read queue to pick up any new entries added concurrently
            queue.RemoveAt(0);
            done.Add(entry);

            // Merge: keep any new entries that were added since we started (entries neither
            // in our working list nor already processed - the file still holds the job that
            // just finished until it is rewritten below, and re-adding it would loop forever)
            foreach (var newEntry in store.Read())
            {
                if (!queue.Any(q => SameJob(q, newEntry)) && !done.Any(d => SameJob(d, newEntry)))
                    queue.Add(newEntry);
            }

            // Write updated queue back (removes completed job, preserves new additions)
            store.Write(queue);
        }

        _ui.Write("");
        WriteSeparator();
        _ui.Write($"QUEUE COMPLETE: {completedCount} job(s) processed successfully!", MessageKind.Success);
        WriteSeparator();
        _ui.SetStatus("RipDisc - Queue Complete");
        _ui.Write("");

        return RipResult.Success;
    }

    private static bool SameJob(QueueEntry a, QueueEntry b) =>
        a.Title == b.Title && a.Disc == b.Disc && a.QueuedAt == b.QueuedAt;

    private void WriteSeparator() => _ui.Write("========================================", MessageKind.Header);
}

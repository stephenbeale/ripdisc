namespace RipDisc.Core;

public class ProcessingStep
{
    public int Number { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}

public class StepTracker
{
    private readonly List<ProcessingStep> _allSteps;
    private readonly List<ProcessingStep> _completedSteps;
    private ProcessingStep? _currentStep;

    public StepTracker()
    {
        _allSteps = new List<ProcessingStep>
        {
            new() { Number = 1, Name = "MakeMKV rip", Description = "Rip disc to MKV files" },
            new() { Number = 2, Name = "HandBrake encoding", Description = "Encode MKV to MP4" },
            new() { Number = 3, Name = "Organize files", Description = "Rename and move files" },
            new() { Number = 4, Name = "Open directory", Description = "Open output folder" }
        };
        _completedSteps = new List<ProcessingStep>();
    }

    public void SetCurrentStep(int stepNumber)
    {
        _currentStep = _allSteps.FirstOrDefault(s => s.Number == stepNumber);
    }

    public void CompleteCurrentStep()
    {
        if (_currentStep != null && !_completedSteps.Contains(_currentStep))
        {
            _completedSteps.Add(_currentStep);
        }
    }

    public List<ProcessingStep> GetRemainingSteps()
    {
        var completedNumbers = _completedSteps.Select(s => s.Number).ToHashSet();
        return _allSteps.Where(s => !completedNumbers.Contains(s.Number)).ToList();
    }

    public List<ProcessingStep> CompletedSteps => _completedSteps;

    public void WriteSummary(IRipUI ui, bool showRemaining = false)
    {
        ui.Write("");
        ui.Write("--- STEPS COMPLETED ---", MessageKind.Success);
        if (_completedSteps.Count == 0)
        {
            ui.Write("  (none)", MessageKind.Detail);
        }
        else
        {
            foreach (var step in _completedSteps)
            {
                ui.Write($"  [X] Step {step.Number}/4: {step.Name}", MessageKind.Success);
            }
        }

        if (showRemaining)
        {
            var remaining = GetRemainingSteps();
            if (remaining.Count > 0)
            {
                ui.Write("");
                ui.Write("--- STEPS REMAINING ---", MessageKind.Warning);
                foreach (var step in remaining)
                {
                    ui.Write($"  [ ] Step {step.Number}/4: {step.Name} - {step.Description}", MessageKind.Warning);
                }
            }
        }
    }
}

namespace RipDisc.Tests;

// Console.SetIn/SetOut are process-wide, so these must not run alongside other console tests
[Collection("Console")]
public class ConsoleRipUITests
{
    private static T WithInput<T>(string input, Func<ConsoleRipUI, T> action)
    {
        var originalIn = Console.In;
        var originalOut = Console.Out;
        try
        {
            Console.SetIn(new StringReader(input));
            Console.SetOut(new StringWriter());
            return action(new ConsoleRipUI());
        }
        finally
        {
            Console.SetIn(originalIn);
            Console.SetOut(originalOut);
        }
    }

    [Fact]
    public void Confirm_ClosedInput_IsNeverConsent()
    {
        Assert.False(WithInput("", ui => ui.Confirm("Start?", defaultAnswer: true)));
    }

    [Fact]
    public void Confirm_BlankLine_TakesDefault()
    {
        Assert.True(WithInput("\n", ui => ui.Confirm("Start?", defaultAnswer: true)));
        Assert.False(WithInput("\n", ui => ui.Confirm("Start?", defaultAnswer: false)));
    }

    [Fact]
    public void Confirm_RepromptsOnNonsense()
    {
        Assert.True(WithInput("maybe\nyes\n", ui => ui.Confirm("Start?", defaultAnswer: false)));
    }

    [Fact]
    public void Choose_ReturnsZeroBasedIndex_AfterInvalidInput()
    {
        Assert.Equal(1, WithInput("9\nabc\n2\n", ui => ui.Choose("Pick", new[] { "a", "b" })));
    }

    [Fact]
    public void Choose_ClosedInput_Cancels()
    {
        Assert.Equal(-1, WithInput("", ui => ui.Choose("Pick", new[] { "a", "b" })));
    }
}

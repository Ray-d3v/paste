namespace PasteWinUI.Tests;

public sealed class PasteRiskEvaluatorTests
{
    [Fact]
    public void TryGetPasteRiskReason_ForLargeText_ReturnsLengthWarning()
    {
        var entry = TestEntryFactory.Create(content: new string('a', 500));

        var actual = PasteRiskEvaluator.TryGetPasteRiskReason(entry, 500, out var reason);

        Assert.True(actual);
        Assert.Contains("Large paste detected", reason, StringComparison.Ordinal);
    }

    [Fact]
    public void TryGetPasteRiskReason_ForMultilineText_ReturnsMultilineWarning()
    {
        var entry = TestEntryFactory.Create(content: "line1\nline2");

        var actual = PasteRiskEvaluator.TryGetPasteRiskReason(entry, 500, out var reason);

        Assert.True(actual);
        Assert.Equal("Multi-line content detected.", reason);
    }

    [Theory]
    [InlineData("person@example.com")]
    [InlineData("ghp_12345678901234567890")]
    public void TryGetPasteRiskReason_ForSensitiveLookingContent_ReturnsWarning(string content)
    {
        var entry = TestEntryFactory.Create(content: content);

        var actual = PasteRiskEvaluator.TryGetPasteRiskReason(entry, 500, out var reason);

        Assert.True(actual);
        Assert.Equal("Sensitive-looking content detected.", reason);
    }

    [Fact]
    public void TryGetPasteRiskReason_ForNonTextContent_ReturnsFalse()
    {
        var entry = TestEntryFactory.Create(kind: "Image", content: "[Image copied]");

        var actual = PasteRiskEvaluator.TryGetPasteRiskReason(entry, 500, out var reason);

        Assert.False(actual);
        Assert.Equal(string.Empty, reason);
    }

    [Fact]
    public void TryGetPasteRiskReason_ForEmptyText_ReturnsFalse()
    {
        var entry = TestEntryFactory.Create(content: "");

        var actual = PasteRiskEvaluator.TryGetPasteRiskReason(entry, 500, out var reason);

        Assert.False(actual);
        Assert.Equal(string.Empty, reason);
    }
}

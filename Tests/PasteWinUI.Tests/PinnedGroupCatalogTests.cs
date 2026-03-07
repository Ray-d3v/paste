using Windows.UI;

namespace PasteWinUI.Tests;

public sealed class PinnedGroupCatalogTests
{
    [Theory]
    [InlineData(null, null)]
    [InlineData("", null)]
    [InlineData("   ", null)]
    [InlineData("Quick", "quick")]
    [InlineData(" WORK ", "work")]
    [InlineData("Idea", "idea")]
    [InlineData("archive", null)]
    public void Normalize_WhenCalled_ReturnsExpectedValue(string? input, string? expected)
    {
        var actual = PinnedGroupCatalog.Normalize(input);

        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData(null, "All")]
    [InlineData("quick", "Quick")]
    [InlineData("work", "Work")]
    [InlineData("idea", "Idea")]
    [InlineData("unknown", "All")]
    public void GetLabel_WhenCalled_ReturnsExpectedLabel(string? input, string expected)
    {
        var actual = PinnedGroupCatalog.GetLabel(input);

        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("quick", 245, 158, 11)]
    [InlineData("work", 34, 197, 94)]
    [InlineData("idea", 96, 165, 250)]
    [InlineData(null, 138, 160, 175)]
    public void GetColor_WhenCalled_ReturnsStableColor(string? input, byte r, byte g, byte b)
    {
        var actual = PinnedGroupCatalog.GetColor(input);

        Assert.Equal(Color.FromArgb(255, r, g, b), actual);
    }
}

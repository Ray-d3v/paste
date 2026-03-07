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
    [InlineData("quick", 242, 183, 102)]
    [InlineData("work", 97, 211, 166)]
    [InlineData("idea", 125, 183, 255)]
    [InlineData(null, 126, 144, 157)]
    public void GetColor_WhenCalled_ReturnsStableColor(string? input, byte r, byte g, byte b)
    {
        var actual = PinnedGroupCatalog.GetColor(input);

        Assert.Equal(Color.FromArgb(255, r, g, b), actual);
    }
}

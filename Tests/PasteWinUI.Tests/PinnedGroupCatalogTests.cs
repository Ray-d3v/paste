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
    [InlineData("archive", "archive")]
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
    [InlineData("quick", 196, 160, 107)]
    [InlineData("work", 127, 181, 157)]
    [InlineData("idea", 136, 169, 201)]
    [InlineData(null, 122, 133, 143)]
    public void GetColor_WhenCalled_ReturnsStableColor(string? input, byte r, byte g, byte b)
    {
        var actual = PinnedGroupCatalog.GetColor(input);

        Assert.Equal(Color.FromArgb(255, r, g, b), actual);
    }

    [Fact]
    public void CreateGroup_WhenNameCollides_GeneratesUniqueStableId()
    {
        var groups = PinnedGroupCatalog.CreateSeededGroups();

        var created = PinnedGroupCatalog.CreateGroup("Quick", groups);

        Assert.Equal("quick-2", created.Id);
        Assert.Equal("Quick", created.Name);
    }

    [Fact]
    public void RenameGroup_WhenCalled_KeepsIdAndUpdatesName()
    {
        var original = new PinnedGroupDefinitionModel("ideas", "Ideas", "blue", 3);

        var renamed = PinnedGroupCatalog.RenameGroup(original, "Research");

        Assert.Equal("ideas", renamed.Id);
        Assert.Equal("Research", renamed.Name);
        Assert.Equal("blue", renamed.ColorKey);
    }

    [Fact]
    public void TryValidateName_WhenDuplicateIgnoringCase_ReturnsFalse()
    {
        var groups = PinnedGroupCatalog.CreateSeededGroups();

        var valid = PinnedGroupCatalog.TryValidateName("quick", groups, null, out _, out var error);

        Assert.False(valid);
        Assert.Equal("A group with that name already exists.", error);
    }
}

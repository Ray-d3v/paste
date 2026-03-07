namespace PasteWinUI.Tests;

public sealed class ClipboardHistoryLogicTests
{
    [Theory]
    [InlineData("https://example.com/path/#frag", "https://example.com/path")]
    [InlineData("https://example.com/path///", "https://example.com/path")]
    [InlineData(" https://example.com/ ", "https://example.com")]
    [InlineData("not-a-url", "not-a-url")]
    public void NormalizeLinkForComparison_WhenCalled_ReturnsExpectedValue(string rawLink, string expected)
    {
        var actual = ClipboardHistoryLogic.NormalizeLinkForComparison(rawLink);

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void IsLikelyDuplicateClipboardEntry_ForTextWithinWindow_ReturnsTrue()
    {
        var copiedAt = new DateTimeOffset(2026, 3, 7, 12, 0, 0, TimeSpan.Zero);
        var existing = TestEntryFactory.Create(content: "same", copiedAtUtc: copiedAt);
        var incoming = TestEntryFactory.Create(content: "same", copiedAtUtc: copiedAt.AddSeconds(1));

        var actual = ClipboardHistoryLogic.IsLikelyDuplicateClipboardEntry(existing, incoming, TimeSpan.FromSeconds(2));

        Assert.True(actual);
    }

    [Fact]
    public void IsLikelyDuplicateClipboardEntry_ForTextOutsideWindow_ReturnsFalse()
    {
        var copiedAt = new DateTimeOffset(2026, 3, 7, 12, 0, 0, TimeSpan.Zero);
        var existing = TestEntryFactory.Create(content: "same", copiedAtUtc: copiedAt);
        var incoming = TestEntryFactory.Create(content: "same", copiedAtUtc: copiedAt.AddSeconds(5));

        var actual = ClipboardHistoryLogic.IsLikelyDuplicateClipboardEntry(existing, incoming, TimeSpan.FromSeconds(2));

        Assert.False(actual);
    }

    [Fact]
    public void IsLikelyDuplicateClipboardEntry_ForLinks_UsesNormalizedUrl()
    {
        var existing = TestEntryFactory.Create(kind: "Link", content: "[Link]", linkUrl: "https://example.com/path/");
        var incoming = TestEntryFactory.Create(
            kind: "Link",
            content: "[Link]",
            linkUrl: "https://example.com/path/#fragment",
            copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 0, 1, TimeSpan.Zero));

        var actual = ClipboardHistoryLogic.IsLikelyDuplicateClipboardEntry(existing, incoming, TimeSpan.FromSeconds(2));

        Assert.True(actual);
    }

    [Fact]
    public void IsLikelyDuplicateClipboardEntry_ForImages_RequiresDimensionsAndBytes()
    {
        var existing = TestEntryFactory.Create(kind: "Image", imageBytes: [1, 2, 3], imageWidth: 320, imageHeight: 180);
        var incoming = TestEntryFactory.Create(
            kind: "Image",
            imageBytes: [9, 8, 7],
            imageWidth: 320,
            imageHeight: 180,
            copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 0, 1, TimeSpan.Zero));

        var actual = ClipboardHistoryLogic.IsLikelyDuplicateClipboardEntry(existing, incoming, TimeSpan.FromSeconds(2));

        Assert.True(actual);
    }

    [Fact]
    public void MergeClipboardEntries_PrefersRicherIncomingMetadata_AndPreservesState()
    {
        var existing = TestEntryFactory.Create(
            kind: "Link",
            content: "[Link]",
            copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 0, 0, TimeSpan.Zero),
            linkUrl: "https://example.com/old",
            linkTitle: null,
            linkHost: null,
            isPinned: true,
            pinnedGroupId: "quick",
            isDeleted: true,
            deletedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 1, 0, TimeSpan.Zero));
        var incoming = TestEntryFactory.Create(
            kind: "Link",
            content: "https://example.com/new",
            copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 2, 0, TimeSpan.Zero),
            linkUrl: "https://example.com/new",
            linkTitle: "Example",
            linkHost: "example.com");

        var merged = ClipboardHistoryLogic.MergeClipboardEntries(existing, incoming);

        Assert.Equal("https://example.com/new", merged.Content);
        Assert.Equal("https://example.com/new", merged.LinkUrl);
        Assert.Equal("Example", merged.LinkTitle);
        Assert.Equal("example.com", merged.LinkHost);
        Assert.Equal(incoming.CopiedAtUtc, merged.CopiedAtUtc);
        Assert.True(merged.IsPinned);
        Assert.Equal("quick", merged.PinnedGroupId);
        Assert.True(merged.IsDeleted);
        Assert.Equal(existing.DeletedAtUtc, merged.DeletedAtUtc);
    }

    [Fact]
    public void CloneEntry_CanPinUnpinTrashAndRestore()
    {
        var original = TestEntryFactory.Create(isPinned: false, isDeleted: false);

        var pinned = ClipboardHistoryLogic.CloneEntry(original, pinnedGroupId: "work", overwritePinnedGroup: true);
        var unpinned = ClipboardHistoryLogic.CloneEntry(pinned, isPinned: false);
        var trashedAt = new DateTimeOffset(2026, 3, 7, 12, 5, 0, TimeSpan.Zero);
        var trashed = ClipboardHistoryLogic.CloneEntry(unpinned, isDeleted: true, deletedAtUtc: trashedAt);
        var restored = ClipboardHistoryLogic.CloneEntry(trashed, isDeleted: false, deletedAtUtc: null);

        Assert.True(pinned.IsPinned);
        Assert.Equal("work", pinned.PinnedGroupId);
        Assert.False(unpinned.IsPinned);
        Assert.Null(unpinned.PinnedGroupId);
        Assert.True(trashed.IsDeleted);
        Assert.Equal(trashedAt, trashed.DeletedAtUtc);
        Assert.False(restored.IsDeleted);
        Assert.Null(restored.DeletedAtUtc);
    }
}

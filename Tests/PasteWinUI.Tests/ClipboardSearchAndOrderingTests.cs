namespace PasteWinUI.Tests;

public sealed class ClipboardSearchAndOrderingTests
{
    [Theory]
    [InlineData("text", "Text", "sample", null, null, null)]
    [InlineData("vscode", "Text", "sample", null, null, null)]
    [InlineData("sample", "Text", "sample", null, null, null)]
    [InlineData("docs", "Link", "[Link]", "Project Docs", "https://example.com/docs", "example.com")]
    [InlineData("example.com", "Link", "[Link]", "Project Docs", "https://example.com/docs", "example.com")]
    public void EntryMatchesSearch_WhenFieldContainsQuery_ReturnsTrue(
        string query,
        string kind,
        string content,
        string? linkTitle,
        string? linkUrl,
        string? linkHost)
    {
        var entry = TestEntryFactory.Create(
            kind: kind,
            content: content,
            linkTitle: linkTitle,
            linkUrl: linkUrl,
            linkHost: linkHost);

        var actual = ClipboardHistoryLogic.EntryMatchesSearch(entry, query);

        Assert.True(actual);
    }

    [Fact]
    public void GetVisibleEntries_InHistoryView_ExcludesDeletedEntries()
    {
        var active = TestEntryFactory.Create(content: "active", copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 0, 0, TimeSpan.Zero));
        var deleted = TestEntryFactory.Create(content: "deleted", isDeleted: true, copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 1, 0, TimeSpan.Zero));

        var entries = ClipboardHistoryLogic.GetVisibleEntries([active, deleted], isTrashView: false, activePinnedGroupId: null, searchQuery: string.Empty);

        Assert.Single(entries);
        Assert.Equal("active", entries[0].Content);
    }

    [Fact]
    public void GetVisibleEntries_InTrashView_IncludesOnlyDeletedEntries()
    {
        var active = TestEntryFactory.Create(content: "active");
        var deleted = TestEntryFactory.Create(content: "deleted", isDeleted: true);

        var entries = ClipboardHistoryLogic.GetVisibleEntries([active, deleted], isTrashView: true, activePinnedGroupId: null, searchQuery: string.Empty);

        Assert.Single(entries);
        Assert.Equal("deleted", entries[0].Content);
    }

    [Fact]
    public void GetVisibleEntries_WithPinnedGroupFilter_NarrowsResults()
    {
        var quick = TestEntryFactory.Create(content: "quick", isPinned: true, pinnedGroupId: "quick");
        var work = TestEntryFactory.Create(content: "work", isPinned: true, pinnedGroupId: "work");

        var entries = ClipboardHistoryLogic.GetVisibleEntries([quick, work], isTrashView: false, activePinnedGroupId: "work", searchQuery: string.Empty);

        Assert.Single(entries);
        Assert.Equal("work", entries[0].Content);
    }

    [Fact]
    public void GetVisibleEntries_DefaultOrdering_IsPinnedFirstThenNewest()
    {
        var newestUnpinned = TestEntryFactory.Create(content: "newest-unpinned", copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 3, 0, TimeSpan.Zero));
        var oldestPinned = TestEntryFactory.Create(
            content: "oldest-pinned",
            copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 0, 0, TimeSpan.Zero),
            isPinned: true,
            pinnedGroupId: "quick");
        var newestPinned = TestEntryFactory.Create(
            content: "newest-pinned",
            copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 2, 0, TimeSpan.Zero),
            isPinned: true,
            pinnedGroupId: "work");

        var entries = ClipboardHistoryLogic.GetVisibleEntries(
            [newestUnpinned, oldestPinned, newestPinned],
            isTrashView: false,
            activePinnedGroupId: null,
            searchQuery: string.Empty);

        Assert.Equal(["newest-pinned", "oldest-pinned", "newest-unpinned"], entries.Select(entry => entry.Content).ToArray());
    }

    [Fact]
    public void GetVisibleEntries_SelectedPinnedGroup_UsesNewestFirstOrdering()
    {
        var olderWork = TestEntryFactory.Create(
            content: "older-work",
            copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 0, 0, TimeSpan.Zero),
            isPinned: true,
            pinnedGroupId: "work");
        var newerWork = TestEntryFactory.Create(
            content: "newer-work",
            copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 3, 0, TimeSpan.Zero),
            isPinned: true,
            pinnedGroupId: "work");
        var pinnedOther = TestEntryFactory.Create(
            content: "quick",
            copiedAtUtc: new DateTimeOffset(2026, 3, 7, 12, 4, 0, TimeSpan.Zero),
            isPinned: true,
            pinnedGroupId: "quick");

        var entries = ClipboardHistoryLogic.GetVisibleEntries(
            [olderWork, newerWork, pinnedOther],
            isTrashView: false,
            activePinnedGroupId: "work",
            searchQuery: string.Empty);

        Assert.Equal(["newer-work", "older-work"], entries.Select(entry => entry.Content).ToArray());
    }
}

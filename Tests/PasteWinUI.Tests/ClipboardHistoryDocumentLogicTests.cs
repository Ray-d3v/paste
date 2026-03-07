namespace PasteWinUI.Tests;

public sealed class ClipboardHistoryDocumentLogicTests
{
    [Fact]
    public void CreateMigratedDocumentFromLegacyEntries_SeedsDefaultGroups()
    {
        var entry = TestEntryFactory.Create(isPinned: true, pinnedGroupId: "quick");

        var document = ClipboardHistoryDocumentLogic.CreateMigratedDocumentFromLegacyEntries([entry]);

        Assert.Equal(["quick", "work", "idea"], document.Groups.Select(group => group.Id).ToArray());
        Assert.Single(document.Entries);
        Assert.Equal("quick", document.Entries[0].PinnedGroupId);
    }

    [Fact]
    public void CreateDocument_WhenEntryReferencesUnknownGroup_UnassignsEntry()
    {
        PinnedGroupDefinitionModel[] groups = [new PinnedGroupDefinitionModel("custom", "Custom", "amber", 0)];
        var entry = TestEntryFactory.Create(isPinned: true, pinnedGroupId: "missing");

        var document = ClipboardHistoryDocumentLogic.CreateDocument(groups, [entry]);

        Assert.Single(document.Entries);
        Assert.Null(document.Entries[0].PinnedGroupId);
        Assert.False(document.Entries[0].IsPinned);
    }

    [Fact]
    public void DeleteGroup_RemovesGroupAndPreservesEntries()
    {
        var groups = PinnedGroupCatalog.CreateSeededGroups();
        var entry = TestEntryFactory.Create(content: "kept", isPinned: true, pinnedGroupId: "work");

        var document = ClipboardHistoryDocumentLogic.DeleteGroup(groups, [entry], "work");

        Assert.DoesNotContain(document.Groups, group => group.Id == "work");
        Assert.Single(document.Entries);
        Assert.Equal("kept", document.Entries[0].Content);
        Assert.Null(document.Entries[0].PinnedGroupId);
        Assert.False(document.Entries[0].IsPinned);
    }
}

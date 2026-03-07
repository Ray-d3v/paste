namespace PasteWinUI;

internal sealed class ClipboardHistoryDocumentModel
{
    public ClipboardHistoryDocumentModel(
        IReadOnlyList<PinnedGroupDefinitionModel> groups,
        IReadOnlyList<ClipboardEntryModel> entries)
    {
        Groups = groups;
        Entries = entries;
    }

    public IReadOnlyList<PinnedGroupDefinitionModel> Groups { get; }
    public IReadOnlyList<ClipboardEntryModel> Entries { get; }
}

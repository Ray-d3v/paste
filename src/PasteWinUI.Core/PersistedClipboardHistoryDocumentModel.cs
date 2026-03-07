namespace PasteWinUI;

internal sealed class PersistedClipboardHistoryDocumentModel
{
    public List<PersistedPinnedGroupDefinitionModel>? Groups { get; set; }
    public List<PersistedClipboardEntryModel>? Entries { get; set; }
}

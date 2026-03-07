namespace PasteWinUI;

internal static class ClipboardHistoryDocumentLogic
{
    internal static ClipboardHistoryDocumentModel CreateMigratedDocumentFromLegacyEntries(
        IEnumerable<ClipboardEntryModel> entries)
    {
        return CreateDocument(PinnedGroupCatalog.CreateSeededGroups(), entries);
    }

    internal static ClipboardHistoryDocumentModel CreateDocument(
        IEnumerable<PinnedGroupDefinitionModel> groups,
        IEnumerable<ClipboardEntryModel> entries)
    {
        var normalizedGroups = PinnedGroupCatalog.SanitizeGroups(groups).ToList();
        var validGroupIds = normalizedGroups
            .Select(group => group.Id)
            .ToHashSet(StringComparer.Ordinal);

        var normalizedEntries = entries
            .Select(entry =>
            {
                if (string.IsNullOrWhiteSpace(entry.PinnedGroupId) || validGroupIds.Contains(entry.PinnedGroupId))
                {
                    return entry;
                }

                return ClipboardHistoryLogic.CloneEntry(entry, pinnedGroupId: null, overwritePinnedGroup: true);
            })
            .OrderByDescending(entry => entry.CopiedAtUtc)
            .ToList();

        return new ClipboardHistoryDocumentModel(normalizedGroups, normalizedEntries);
    }

    internal static ClipboardHistoryDocumentModel DeleteGroup(
        IEnumerable<PinnedGroupDefinitionModel> groups,
        IEnumerable<ClipboardEntryModel> entries,
        string groupId)
    {
        var normalizedId = PinnedGroupCatalog.Normalize(groupId);
        if (string.IsNullOrWhiteSpace(normalizedId))
        {
            return CreateDocument(groups, entries);
        }

        var nextGroups = groups
            .Where(group => !string.Equals(group.Id, normalizedId, StringComparison.Ordinal))
            .ToList();
        var nextEntries = entries
            .Select(entry => string.Equals(entry.PinnedGroupId, normalizedId, StringComparison.Ordinal)
                ? ClipboardHistoryLogic.CloneEntry(entry, pinnedGroupId: null, overwritePinnedGroup: true)
                : entry)
            .ToList();

        return CreateDocument(nextGroups, nextEntries);
    }
}

namespace PasteWinUI;

internal static class ClipboardHistoryLogic
{
    internal static bool IsLikelyDuplicateClipboardEntry(
        ClipboardEntryModel existing,
        ClipboardEntryModel incoming,
        TimeSpan duplicateMergeWindow)
    {
        var timeDiff = (incoming.CopiedAtUtc - existing.CopiedAtUtc).Duration();
        if (timeDiff > duplicateMergeWindow)
        {
            return false;
        }

        if (!string.Equals(existing.Kind, incoming.Kind, StringComparison.Ordinal))
        {
            return false;
        }

        if (string.Equals(incoming.Kind, "Link", StringComparison.Ordinal))
        {
            var left = NormalizeLinkForComparison(existing.LinkUrl ?? existing.Content);
            var right = NormalizeLinkForComparison(incoming.LinkUrl ?? incoming.Content);
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }

        if (string.Equals(incoming.Kind, "Image", StringComparison.Ordinal))
        {
            return existing.ImageWidth == incoming.ImageWidth &&
                   existing.ImageHeight == incoming.ImageHeight &&
                   existing.ImagePngBytes is { Length: > 0 } &&
                   incoming.ImagePngBytes is { Length: > 0 };
        }

        return string.Equals(existing.Content, incoming.Content, StringComparison.Ordinal);
    }

    internal static ClipboardEntryModel MergeClipboardEntries(ClipboardEntryModel existing, ClipboardEntryModel incoming)
    {
        var content = ChoosePreferredContent(incoming.Content, existing.Content);
        var sourceExePath = incoming.SourceExePath ?? existing.SourceExePath;
        var sourceIconPngBytes = incoming.SourceIconPngBytes ?? existing.SourceIconPngBytes;
        var sourceHeaderColor = incoming.SourceHeaderColor;
        var imagePngBytes = incoming.ImagePngBytes ?? existing.ImagePngBytes;
        var imageWidth = incoming.ImageWidth > 0 ? incoming.ImageWidth : existing.ImageWidth;
        var imageHeight = incoming.ImageHeight > 0 ? incoming.ImageHeight : existing.ImageHeight;
        var linkUrl = ChoosePreferredText(incoming.LinkUrl, existing.LinkUrl);
        var linkTitle = ChoosePreferredText(incoming.LinkTitle, existing.LinkTitle);
        var linkPreviewImageBytes = incoming.LinkPreviewImageBytes ?? existing.LinkPreviewImageBytes;
        var linkFaviconImageBytes = incoming.LinkFaviconImageBytes ?? existing.LinkFaviconImageBytes;
        var linkHost = ChoosePreferredText(incoming.LinkHost, existing.LinkHost);

        return new ClipboardEntryModel(
            incoming.Kind,
            incoming.SourceApp,
            content,
            incoming.CopiedAtUtc > existing.CopiedAtUtc ? incoming.CopiedAtUtc : existing.CopiedAtUtc,
            sourceExePath,
            sourceIconPngBytes,
            sourceHeaderColor,
            imagePngBytes,
            imageWidth,
            imageHeight,
            linkUrl,
            linkTitle,
            linkPreviewImageBytes,
            linkFaviconImageBytes,
            linkHost,
            existing.IsPinned || incoming.IsPinned,
            existing.IsDeleted || incoming.IsDeleted,
            existing.DeletedAtUtc ?? incoming.DeletedAtUtc,
            existing.PinnedGroupId ?? incoming.PinnedGroupId);
    }

    internal static ClipboardEntryModel CloneEntry(
        ClipboardEntryModel entry,
        bool? isPinned = null,
        string? pinnedGroupId = null,
        bool overwritePinnedGroup = false,
        bool? isDeleted = null,
        DateTimeOffset? deletedAtUtc = null)
    {
        var resolvedPinnedGroupId = overwritePinnedGroup
            ? PinnedGroupCatalog.Normalize(pinnedGroupId)
            : (isPinned.HasValue
                ? (isPinned.Value ? (entry.PinnedGroupId ?? PinnedGroupCatalog.Quick) : null)
                : entry.PinnedGroupId);
        var resolvedDeletedAtUtc = isDeleted == false
            ? null
            : (deletedAtUtc.HasValue ? deletedAtUtc : entry.DeletedAtUtc);

        return new ClipboardEntryModel(
            entry.Kind,
            entry.SourceApp,
            entry.Content,
            entry.CopiedAtUtc,
            entry.SourceExePath,
            entry.SourceIconPngBytes,
            entry.SourceHeaderColor,
            entry.ImagePngBytes,
            entry.ImageWidth,
            entry.ImageHeight,
            entry.LinkUrl,
            entry.LinkTitle,
            entry.LinkPreviewImageBytes,
            entry.LinkFaviconImageBytes,
            entry.LinkHost,
            !string.IsNullOrWhiteSpace(resolvedPinnedGroupId),
            isDeleted ?? entry.IsDeleted,
            resolvedDeletedAtUtc,
            resolvedPinnedGroupId);
    }

    internal static string NormalizeLinkForComparison(string? rawLink)
    {
        if (string.IsNullOrWhiteSpace(rawLink))
        {
            return string.Empty;
        }

        var trimmed = rawLink.Trim();
        if (!Uri.TryCreate(trimmed, UriKind.Absolute, out var uri))
        {
            return trimmed;
        }

        var builder = new UriBuilder(uri)
        {
            Fragment = string.Empty
        };

        if (builder.Path.Length > 1)
        {
            builder.Path = builder.Path.TrimEnd('/');
        }

        var normalized = builder.Uri.AbsoluteUri;
        return normalized.Length > 1 ? normalized.TrimEnd('/') : normalized;
    }

    internal static bool EntryMatchesSearch(ClipboardEntryModel entry, string query)
    {
        var comparison = StringComparison.OrdinalIgnoreCase;
        if (entry.Kind.Contains(query, comparison) ||
            entry.SourceApp.Contains(query, comparison) ||
            entry.Content.Contains(query, comparison))
        {
            return true;
        }

        return (!string.IsNullOrWhiteSpace(entry.LinkTitle) && entry.LinkTitle.Contains(query, comparison)) ||
            (!string.IsNullOrWhiteSpace(entry.LinkUrl) && entry.LinkUrl.Contains(query, comparison)) ||
            (!string.IsNullOrWhiteSpace(entry.LinkHost) && entry.LinkHost.Contains(query, comparison));
    }

    internal static List<ClipboardEntryModel> GetVisibleEntries(
        IEnumerable<ClipboardEntryModel> entries,
        bool isTrashView,
        string? activePinnedGroupId,
        string searchQuery)
    {
        IEnumerable<ClipboardEntryModel> visibleEntries = entries
            .Where(entry => isTrashView ? entry.IsDeleted : !entry.IsDeleted);

        var normalizedPinnedGroupId = PinnedGroupCatalog.Normalize(activePinnedGroupId);
        if (!string.IsNullOrWhiteSpace(normalizedPinnedGroupId))
        {
            visibleEntries = visibleEntries.Where(entry =>
                string.Equals(entry.PinnedGroupId, normalizedPinnedGroupId, StringComparison.Ordinal));
        }

        var orderedEntries = string.IsNullOrWhiteSpace(normalizedPinnedGroupId)
            ? visibleEntries
                .OrderByDescending(entry => !string.IsNullOrWhiteSpace(entry.PinnedGroupId))
                .ThenByDescending(entry => entry.CopiedAtUtc)
            : visibleEntries.OrderByDescending(entry => entry.CopiedAtUtc);

        if (string.IsNullOrWhiteSpace(searchQuery))
        {
            return orderedEntries.ToList();
        }

        return orderedEntries
            .Where(entry => EntryMatchesSearch(entry, searchQuery))
            .ToList();
    }

    private static string ChoosePreferredContent(string preferred, string fallback)
    {
        if (string.IsNullOrWhiteSpace(preferred))
        {
            return fallback;
        }

        if (preferred is "[Link]" or "[Image copied]" && !string.IsNullOrWhiteSpace(fallback))
        {
            return fallback;
        }

        return preferred;
    }

    private static string? ChoosePreferredText(string? preferred, string? fallback)
    {
        return string.IsNullOrWhiteSpace(preferred) ? fallback : preferred;
    }
}

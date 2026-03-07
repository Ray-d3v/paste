using Windows.UI;

namespace PasteWinUI.Tests;

internal static class TestEntryFactory
{
    internal static ClipboardEntryModel Create(
        string kind = "Text",
        string sourceApp = "vscode",
        string content = "sample",
        DateTimeOffset? copiedAtUtc = null,
        string? linkUrl = null,
        string? linkTitle = null,
        string? linkHost = null,
        byte[]? imageBytes = null,
        int imageWidth = 0,
        int imageHeight = 0,
        bool isPinned = false,
        string? pinnedGroupId = null,
        bool isDeleted = false,
        DateTimeOffset? deletedAtUtc = null)
    {
        return new ClipboardEntryModel(
            kind,
            sourceApp,
            content,
            copiedAtUtc ?? new DateTimeOffset(2026, 3, 7, 12, 0, 0, TimeSpan.Zero),
            sourceExePath: null,
            sourceIconPngBytes: null,
            sourceHeaderColor: Color.FromArgb(255, 10, 20, 30),
            imagePngBytes: imageBytes,
            imageWidth: imageWidth,
            imageHeight: imageHeight,
            linkUrl: linkUrl,
            linkTitle: linkTitle,
            linkPreviewImageBytes: null,
            linkFaviconImageBytes: null,
            linkHost: linkHost,
            isPinned: isPinned,
            isDeleted: isDeleted,
            deletedAtUtc: deletedAtUtc,
            pinnedGroupId: pinnedGroupId);
    }
}

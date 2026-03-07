using Windows.UI;

namespace PasteWinUI;

internal sealed class ClipboardEntryModel
{
    public ClipboardEntryModel(
        string kind,
        string sourceApp,
        string content,
        DateTimeOffset copiedAtUtc,
        string? sourceExePath,
        byte[]? sourceIconPngBytes,
        Color sourceHeaderColor,
        byte[]? imagePngBytes = null,
        int imageWidth = 0,
        int imageHeight = 0,
        string? linkUrl = null,
        string? linkTitle = null,
        byte[]? linkPreviewImageBytes = null,
        byte[]? linkFaviconImageBytes = null,
        string? linkHost = null,
        bool isPinned = false,
        bool isDeleted = false,
        DateTimeOffset? deletedAtUtc = null,
        string? pinnedGroupId = null)
    {
        Kind = kind;
        SourceApp = sourceApp;
        Content = content;
        CopiedAtUtc = copiedAtUtc;
        SourceExePath = sourceExePath;
        SourceIconPngBytes = sourceIconPngBytes;
        SourceHeaderColor = sourceHeaderColor;
        ImagePngBytes = imagePngBytes;
        ImageWidth = imageWidth;
        ImageHeight = imageHeight;
        LinkUrl = linkUrl;
        LinkTitle = linkTitle;
        LinkPreviewImageBytes = linkPreviewImageBytes;
        LinkFaviconImageBytes = linkFaviconImageBytes;
        LinkHost = linkHost;
        PinnedGroupId = PinnedGroupCatalog.Normalize(pinnedGroupId);
        if (PinnedGroupId is null && isPinned)
        {
            PinnedGroupId = PinnedGroupCatalog.Quick;
        }

        IsPinned = !string.IsNullOrWhiteSpace(PinnedGroupId);
        IsDeleted = isDeleted;
        DeletedAtUtc = deletedAtUtc;
    }

    public string Kind { get; }
    public string SourceApp { get; }
    public string Content { get; }
    public DateTimeOffset CopiedAtUtc { get; }
    public string? SourceExePath { get; }
    public byte[]? SourceIconPngBytes { get; }
    public Color SourceHeaderColor { get; }
    public byte[]? ImagePngBytes { get; }
    public int ImageWidth { get; }
    public int ImageHeight { get; }
    public string? LinkUrl { get; }
    public string? LinkTitle { get; }
    public byte[]? LinkPreviewImageBytes { get; }
    public byte[]? LinkFaviconImageBytes { get; }
    public string? LinkHost { get; }
    public string? PinnedGroupId { get; }
    public bool IsPinned { get; }
    public bool IsDeleted { get; }
    public DateTimeOffset? DeletedAtUtc { get; }
}

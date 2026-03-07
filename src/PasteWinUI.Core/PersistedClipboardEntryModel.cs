namespace PasteWinUI;

internal sealed class PersistedClipboardEntryModel
{
    public string? Kind { get; set; }
    public string? SourceApp { get; set; }
    public string? Content { get; set; }
    public DateTimeOffset CopiedAtUtc { get; set; }
    public string? SourceExePath { get; set; }
    public byte[]? SourceIconPngBytes { get; set; }
    public uint SourceHeaderColorArgb { get; set; }
    public byte[]? ImagePngBytes { get; set; }
    public int ImageWidth { get; set; }
    public int ImageHeight { get; set; }
    public string? LinkUrl { get; set; }
    public string? LinkTitle { get; set; }
    public byte[]? LinkPreviewImageBytes { get; set; }
    public byte[]? LinkFaviconImageBytes { get; set; }
    public string? LinkHost { get; set; }
    public bool IsPinned { get; set; }
    public string? PinnedGroupId { get; set; }
    public bool IsDeleted { get; set; }
    public DateTimeOffset? DeletedAtUtc { get; set; }
}

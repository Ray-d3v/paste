using Microsoft.UI.Xaml.Media;
using Windows.UI;

namespace PasteWinUI;

internal static class UiPalette
{
    internal static Color AccentPrimary => FromHex("#5CAEA7");
    internal static Color AccentPrimaryHover => FromHex("#74BEB8");
    internal static Color AccentPrimaryPressed => FromHex("#4C8D87");
    internal static Color FocusRing => FromHex("#8FC2BE");
    internal static Color Canvas => FromHex("#0B0F12");
    internal static Color OverlaySurface => FromHex("#11161B");
    internal static Color ElevatedSurface => FromHex("#171D23");
    internal static Color PopupSurface => FromHex("#1B2229");
    internal static Color CardSurface => FromHex("#151C23");
    internal static Color CardSurfaceHover => FromHex("#192129");
    internal static Color StrokeSubtle => FromHex("#2A343E");
    internal static Color StrokeStrong => FromHex("#394651");
    internal static Color TextPrimary => FromHex("#F3F4F1");
    internal static Color TextSecondary => FromHex("#B5BEC6");
    internal static Color TextMuted => FromHex("#818D99");
    internal static Color AccentWarning => FromHex("#C7A36A");
    internal static Color AccentDanger => FromHex("#C56F6F");
    internal static Color PinAll => FromHex("#7A858F");
    internal static Color PinQuick => FromHex("#C4A06B");
    internal static Color PinWork => FromHex("#7FB59D");
    internal static Color PinIdea => FromHex("#88A9C9");

    internal static SolidColorBrush Brush(Color color) => new(color);

    internal static Color WithAlpha(Color color, byte alpha) =>
        Color.FromArgb(alpha, color.R, color.G, color.B);

    internal static Color GetPinnedDotColor(string? groupId)
    {
        return PinnedGroupCatalog.Normalize(groupId) switch
        {
            PinnedGroupCatalog.Quick => PinQuick,
            PinnedGroupCatalog.Work => PinWork,
            PinnedGroupCatalog.Idea => PinIdea,
            _ => PinAll
        };
    }

    internal static Color[] LinkPlaceholderPalette =>
    [
        FromHex("#445664"),
        FromHex("#48606A"),
        FromHex("#52606F"),
        FromHex("#5A6374"),
        FromHex("#4B5A65"),
        FromHex("#5E6971")
    ];

    private static Color FromHex(string hex)
    {
        var value = hex.TrimStart('#');
        if (value.Length == 6)
        {
            value = "FF" + value;
        }

        return Color.FromArgb(
            Convert.ToByte(value.Substring(0, 2), 16),
            Convert.ToByte(value.Substring(2, 2), 16),
            Convert.ToByte(value.Substring(4, 2), 16),
            Convert.ToByte(value.Substring(6, 2), 16));
    }
}

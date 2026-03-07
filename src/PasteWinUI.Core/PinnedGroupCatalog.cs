using Windows.UI;

namespace PasteWinUI;

internal static class PinnedGroupCatalog
{
    internal const string Quick = "quick";
    internal const string Work = "work";
    internal const string Idea = "idea";

    internal static string? Normalize(string? groupId)
    {
        if (string.IsNullOrWhiteSpace(groupId))
        {
            return null;
        }

        return groupId.Trim().ToLowerInvariant() switch
        {
            Quick => Quick,
            Work => Work,
            Idea => Idea,
            _ => null
        };
    }

    internal static string GetLabel(string? groupId)
    {
        return Normalize(groupId) switch
        {
            Quick => "Quick",
            Work => "Work",
            Idea => "Idea",
            _ => "All"
        };
    }

    internal static Color GetColor(string? groupId)
    {
        return Normalize(groupId) switch
        {
            Quick => Color.FromArgb(255, 242, 183, 102),
            Work => Color.FromArgb(255, 97, 211, 166),
            Idea => Color.FromArgb(255, 125, 183, 255),
            _ => Color.FromArgb(255, 126, 144, 157)
        };
    }
}

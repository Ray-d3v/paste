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
            Quick => Color.FromArgb(255, 245, 158, 11),
            Work => Color.FromArgb(255, 34, 197, 94),
            Idea => Color.FromArgb(255, 96, 165, 250),
            _ => Color.FromArgb(255, 138, 160, 175)
        };
    }
}

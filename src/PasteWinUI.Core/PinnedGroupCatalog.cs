using Windows.UI;

namespace PasteWinUI;

internal static class PinnedGroupCatalog
{
    internal const int MaxNameLength = 32;
    internal const string Quick = "quick";
    internal const string Work = "work";
    internal const string Idea = "idea";

    private static readonly (string Id, string Name, string ColorKey, int SortOrder)[] SeededDefinitions =
    [
        (Quick, "Quick", "amber", 0),
        (Work, "Work", "sage", 1),
        (Idea, "Idea", "blue", 2)
    ];

    private static readonly (string Key, Color Color)[] Palette =
    [
        ("amber", Color.FromArgb(255, 196, 160, 107)),
        ("sage", Color.FromArgb(255, 127, 181, 157)),
        ("blue", Color.FromArgb(255, 136, 169, 201)),
        ("slate", Color.FromArgb(255, 122, 133, 143)),
        ("mauve", Color.FromArgb(255, 144, 136, 164)),
        ("olive", Color.FromArgb(255, 146, 162, 116))
    ];

    internal static IReadOnlyList<PinnedGroupDefinitionModel> CreateSeededGroups()
    {
        return SeededDefinitions
            .Select(definition => new PinnedGroupDefinitionModel(
                definition.Id,
                definition.Name,
                definition.ColorKey,
                definition.SortOrder))
            .ToList();
    }

    internal static string? Normalize(string? groupId)
    {
        if (string.IsNullOrWhiteSpace(groupId))
        {
            return null;
        }

        var normalized = groupId.Trim().ToLowerInvariant();
        return normalized.All(ch => char.IsLetterOrDigit(ch) || ch == '-') ? normalized : null;
    }

    internal static string GetLabel(string? groupId) => GetLabel(CreateSeededGroups(), groupId);

    internal static string GetLabel(IEnumerable<PinnedGroupDefinitionModel> groups, string? groupId)
    {
        var normalizedId = Normalize(groupId);
        if (string.IsNullOrWhiteSpace(normalizedId))
        {
            return "All";
        }

        var definition = FindById(groups, normalizedId);
        return definition?.Name ?? "All";
    }

    internal static Color GetColor(string? groupId) => GetColor(CreateSeededGroups(), groupId);

    internal static Color GetColor(IEnumerable<PinnedGroupDefinitionModel> groups, string? groupId)
    {
        var normalizedId = Normalize(groupId);
        if (string.IsNullOrWhiteSpace(normalizedId))
        {
            return ResolveColorByKey("slate");
        }

        var definition = FindById(groups, normalizedId);
        return ResolveColorByKey(definition?.ColorKey);
    }

    internal static IReadOnlyList<PinnedGroupDefinitionModel> SanitizeGroups(IEnumerable<PinnedGroupDefinitionModel> groups)
    {
        var normalizedGroups = new List<PinnedGroupDefinitionModel>();
        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        var sortOrder = 0;

        foreach (var group in groups.OrderBy(group => group.SortOrder))
        {
            var normalizedId = Normalize(group.Id);
            var normalizedName = NormalizeName(group.Name);
            if (string.IsNullOrWhiteSpace(normalizedId) ||
                string.IsNullOrWhiteSpace(normalizedName) ||
                !seenIds.Add(normalizedId))
            {
                continue;
            }

            normalizedGroups.Add(new PinnedGroupDefinitionModel(
                normalizedId,
                normalizedName,
                ResolveColorKey(group.ColorKey, normalizedGroups.Count),
                sortOrder++));
        }

        return normalizedGroups;
    }

    internal static bool TryValidateName(
        string? name,
        IEnumerable<PinnedGroupDefinitionModel> groups,
        string? excludeId,
        out string normalizedName,
        out string error)
    {
        normalizedName = NormalizeName(name);
        if (string.IsNullOrWhiteSpace(normalizedName))
        {
            error = "Name is required.";
            return false;
        }

        if (normalizedName.Length > MaxNameLength)
        {
            error = $"Name must be {MaxNameLength} characters or fewer.";
            return false;
        }

        var normalizedExcludeId = Normalize(excludeId);
        var candidateName = normalizedName;
        if (groups.Any(group =>
                !string.Equals(group.Id, normalizedExcludeId, StringComparison.Ordinal) &&
                string.Equals(group.Name, candidateName, StringComparison.OrdinalIgnoreCase)))
        {
            error = "A group with that name already exists.";
            return false;
        }

        error = string.Empty;
        return true;
    }

    internal static PinnedGroupDefinitionModel CreateGroup(string name, IEnumerable<PinnedGroupDefinitionModel> groups)
    {
        var existing = SanitizeGroups(groups).ToList();
        var id = CreateIdFromName(name, existing);
        var normalizedName = NormalizeName(name);
        var colorKey = ResolveColorKey(null, existing.Count);
        var sortOrder = existing.Count == 0 ? 0 : existing.Max(group => group.SortOrder) + 1;
        return new PinnedGroupDefinitionModel(id, normalizedName, colorKey, sortOrder);
    }

    internal static PinnedGroupDefinitionModel RenameGroup(PinnedGroupDefinitionModel group, string name)
    {
        return new PinnedGroupDefinitionModel(group.Id, NormalizeName(name), group.ColorKey, group.SortOrder);
    }

    internal static PinnedGroupDefinitionModel? FindById(IEnumerable<PinnedGroupDefinitionModel> groups, string? groupId)
    {
        var normalizedId = Normalize(groupId);
        if (string.IsNullOrWhiteSpace(normalizedId))
        {
            return null;
        }

        return groups.FirstOrDefault(group => string.Equals(group.Id, normalizedId, StringComparison.Ordinal));
    }

    internal static string CreateIdFromName(string name, IEnumerable<PinnedGroupDefinitionModel> groups)
    {
        var normalizedName = NormalizeName(name);
        var slug = new string(
            normalizedName
                .ToLowerInvariant()
                .Select(ch => char.IsLetterOrDigit(ch) ? ch : '-')
                .ToArray());
        slug = slug.Trim('-');
        while (slug.Contains("--", StringComparison.Ordinal))
        {
            slug = slug.Replace("--", "-", StringComparison.Ordinal);
        }

        if (string.IsNullOrWhiteSpace(slug))
        {
            slug = "group";
        }

        var existingIds = groups
            .Select(group => group.Id)
            .ToHashSet(StringComparer.Ordinal);
        if (!existingIds.Contains(slug))
        {
            return slug;
        }

        var index = 2;
        while (existingIds.Contains($"{slug}-{index}"))
        {
            index++;
        }

        return $"{slug}-{index}";
    }

    internal static Color ResolveColorByKey(string? colorKey)
    {
        var normalizedKey = colorKey?.Trim().ToLowerInvariant();
        foreach (var (key, color) in Palette)
        {
            if (string.Equals(key, normalizedKey, StringComparison.Ordinal))
            {
                return color;
            }
        }

        return Palette[3].Color;
    }

    private static string NormalizeName(string? name)
    {
        var normalized = (name ?? string.Empty).Trim();
        if (normalized.Length > MaxNameLength)
        {
            normalized = normalized[..MaxNameLength].TrimEnd();
        }

        return normalized;
    }

    private static string ResolveColorKey(string? colorKey, int index)
    {
        var normalizedKey = colorKey?.Trim().ToLowerInvariant();
        if (!string.IsNullOrWhiteSpace(normalizedKey) &&
            Palette.Any(item => string.Equals(item.Key, normalizedKey, StringComparison.Ordinal)))
        {
            return normalizedKey;
        }

        return Palette[index % Palette.Length].Key;
    }
}

namespace PasteWinUI;

internal sealed class PinnedGroupDefinitionModel
{
    public PinnedGroupDefinitionModel(string id, string name, string colorKey, int sortOrder)
    {
        Id = id;
        Name = name;
        ColorKey = colorKey;
        SortOrder = sortOrder;
    }

    public string Id { get; }
    public string Name { get; }
    public string ColorKey { get; }
    public int SortOrder { get; }
}

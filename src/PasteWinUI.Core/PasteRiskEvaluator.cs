using System.Text.RegularExpressions;

namespace PasteWinUI;

internal static partial class PasteRiskEvaluator
{
    private static readonly Regex EmailLikeRegex = new(
        @"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex TokenLikeRegex = new(
        @"\b(?:sk-[A-Za-z0-9]{16,}|ghp_[A-Za-z0-9]{20,}|AKIA[0-9A-Z]{16})\b",
        RegexOptions.Compiled);

    internal static bool TryGetPasteRiskReason(
        ClipboardEntryModel entry,
        int riskPasteLengthThreshold,
        out string reason)
    {
        reason = string.Empty;
        if (!string.Equals(entry.Kind, "Text", StringComparison.Ordinal) &&
            !string.Equals(entry.Kind, "Link", StringComparison.Ordinal))
        {
            return false;
        }

        var text = string.Equals(entry.Kind, "Link", StringComparison.Ordinal)
            ? (string.IsNullOrWhiteSpace(entry.LinkUrl) ? entry.Content : entry.LinkUrl!)
            : entry.Content;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        if (text.Length >= riskPasteLengthThreshold)
        {
            reason = $"Large paste detected ({text.Length:N0} characters).";
            return true;
        }

        if (text.Contains('\n'))
        {
            reason = "Multi-line content detected.";
            return true;
        }

        if (EmailLikeRegex.IsMatch(text) || TokenLikeRegex.IsMatch(text))
        {
            reason = "Sensitive-looking content detected.";
            return true;
        }

        return false;
    }
}

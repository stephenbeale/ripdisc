using System.Text.RegularExpressions;

namespace RipDisc.Core;

public static class TitleValidator
{
    /// <summary>
    /// Warnings for a series title that looks like it carries metadata which belongs in
    /// -season/-disc instead (e.g. "Fargo Season 1 Disc 2"). Empty for movies.
    /// </summary>
    public static List<string> GetMisplacedMetadataWarnings(string title, bool series)
    {
        var warnings = new List<string>();
        if (!series)
            return warnings;

        if (Regex.IsMatch(title, @"(?i)\bseries\s*\d"))
            warnings.Add("Contains 'Series N' - use -Season parameter instead");
        if (Regex.IsMatch(title, @"(?i)\bseason\s*\d"))
            warnings.Add("Contains 'Season N' - use -Season parameter instead");
        if (Regex.IsMatch(title, @"(?i)\bdisc\s*\d"))
            warnings.Add("Contains 'Disc N' - use -Disc parameter instead");
        if (Regex.IsMatch(title, @"(?i)\bS\d{1,2}E\d"))
            warnings.Add("Contains episode code (e.g. S01E01) - use -Series -Season instead");

        return warnings;
    }
}

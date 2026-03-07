using System.Collections.Generic;
using Ace7LocalizationFormat.Formats;

namespace Ace7Ed
{
    /// <summary>
    /// Alias for game localization DAT letter keys (used by Launcher and LocalizationEditor).
    /// </summary>
    public static class AceLocalizationConstants
    {
        public static IReadOnlyDictionary<char, string> DatLetters => DatConstants.DatLetters;
    }

    /// <summary>
    /// Column headers for CSV export/import of localization strings.
    /// Order: Variable (key), then A (English) through M (Simplified Chinese).
    /// </summary>
    public static class CsvExportConstants
    {
        public const string VariableColumn = "Variable";

        /// <summary>
        /// Headers for the 14 columns: Variable, then A (English), B (Trad. Chinese), ... M (Simplified Chinese).
        /// </summary>
        public static readonly string[] ColumnHeaders =
        {
            VariableColumn,
            "A (English)",
            "B (Trad. Chinese)",
            "C (French)",
            "D (German)",
            "E (Italian)",
            "F (Japanese)",
            "G (Korean)",
            "H (Euro. Spanish)",
            "I (Latin American Spanish)",
            "J (Polish)",
            "K (Brazilian Portuguese)",
            "L (Russian)",
            "M (Simplified Chinese)"
        };

        public const int LanguageColumnCount = 13;
    }
}

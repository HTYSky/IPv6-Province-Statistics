using System.Text;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

internal static class HeaderText
{
    public static string Normalize(string? value)
    {
        if (value is null)
        {
            return string.Empty;
        }

        string compatibilityNormalized = value.Normalize(NormalizationForm.FormKC);
        var builder = new StringBuilder(compatibilityNormalized.Length);
        foreach (Rune rune in compatibilityNormalized.EnumerateRunes())
        {
            if (Rune.IsWhiteSpace(rune))
            {
                continue;
            }

            builder.Append(rune.Value switch
            {
                '（' => '(',
                '）' => ')',
                _ => rune.ToString(),
            });
        }

        return builder.ToString();
    }

    public static bool Contains(string? value, string expected) =>
        Normalize(value).Contains(Normalize(expected), StringComparison.OrdinalIgnoreCase);
}

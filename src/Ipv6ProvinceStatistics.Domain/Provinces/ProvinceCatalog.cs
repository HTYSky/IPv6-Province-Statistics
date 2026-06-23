using System.Collections.Frozen;
using System.Text;

namespace Ipv6ProvinceStatistics.Domain.Provinces;

public static class ProvinceCatalog
{
    public static IReadOnlyList<Province> All { get; } = Array.AsReadOnly<Province>(
    [
        new("北京"), new("天津"), new("河北"), new("山西"), new("内蒙古"), new("辽宁"),
        new("吉林"), new("黑龙江"), new("上海"), new("江苏"), new("浙江"), new("安徽"),
        new("福建"), new("江西"), new("山东"), new("河南"), new("湖北"), new("湖南"),
        new("广东"), new("广西"), new("海南"), new("重庆"), new("四川"), new("贵州"),
        new("云南"), new("西藏"), new("陕西"), new("甘肃"), new("青海"), new("宁夏"),
        new("新疆"),
    ]);

    private static readonly FrozenDictionary<string, Province> Aliases = CreateAliases();

    public static bool TryResolve(string? value, out Province province)
    {
        if (value is null)
        {
            province = default;
            return false;
        }

        return Aliases.TryGetValue(RemoveWhitespace(value), out province);
    }

    private static FrozenDictionary<string, Province> CreateAliases()
    {
        var aliases = new Dictionary<string, Province>(StringComparer.Ordinal);

        foreach (Province province in All)
        {
            aliases.Add(province.Name, province);
        }

        AddSuffixAliases(
            aliases,
            [
                "河北", "山西", "辽宁", "吉林", "黑龙江", "江苏", "浙江", "安徽", "福建",
                "江西", "山东", "河南", "湖北", "湖南", "广东", "海南", "四川", "贵州",
                "云南", "陕西", "甘肃", "青海",
            ],
            "省");
        AddSuffixAliases(aliases, ["北京", "天津", "上海", "重庆"], "市");
        AddSuffixAliases(aliases, ["内蒙古", "广西", "西藏", "宁夏", "新疆"], "自治区");

        aliases.Add("广西壮族自治区", aliases["广西"]);
        aliases.Add("宁夏回族自治区", aliases["宁夏"]);
        aliases.Add("新疆维吾尔族自治区", aliases["新疆"]);

        return aliases.ToFrozenDictionary(StringComparer.Ordinal);
    }

    private static void AddSuffixAliases(
        IDictionary<string, Province> aliases,
        IEnumerable<string> provinceNames,
        string suffix)
    {
        foreach (string provinceName in provinceNames)
        {
            aliases.Add($"{provinceName}{suffix}", aliases[provinceName]);
        }
    }

    private static string RemoveWhitespace(string value)
    {
        var normalized = new StringBuilder(value.Length);

        foreach (Rune rune in value.EnumerateRunes())
        {
            if (!Rune.IsWhiteSpace(rune))
            {
                normalized.Append(rune);
            }
        }

        return normalized.ToString();
    }
}

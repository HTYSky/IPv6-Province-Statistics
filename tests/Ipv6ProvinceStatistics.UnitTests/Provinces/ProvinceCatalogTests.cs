using Ipv6ProvinceStatistics.Domain.Provinces;

namespace Ipv6ProvinceStatistics.UnitTests.Provinces;

public sealed class ProvinceCatalogTests
{
    [Fact]
    public void AllContainsThe31UniqueProvincesInStandardOrder()
    {
        string[] expectedNames =
        [
            "北京", "天津", "河北", "山西", "内蒙古", "辽宁", "吉林", "黑龙江", "上海", "江苏",
            "浙江", "安徽", "福建", "江西", "山东", "河南", "湖北", "湖南", "广东", "广西",
            "海南", "重庆", "四川", "贵州", "云南", "西藏", "陕西", "甘肃", "青海", "宁夏", "新疆",
        ];

        Assert.Equal(31, ProvinceCatalog.All.Count);
        Assert.Equal(expectedNames, ProvinceCatalog.All.Select(province => province.Name));
        Assert.Equal(31, ProvinceCatalog.All.Select(province => province.Name).Distinct().Count());
    }

    [Theory]
    [InlineData("北京市", "北京")]
    [InlineData("北   京", "北京")]
    [InlineData("北\u3000京", "北京")]
    [InlineData("内蒙古自治区", "内蒙古")]
    [InlineData("广西壮族自治区", "广西")]
    [InlineData("宁夏回族自治区", "宁夏")]
    [InlineData("新疆维吾尔族自治区", "新疆")]
    [InlineData("西藏自治区", "西藏")]
    public void TryResolveRecognizesSupportedAliases(string value, string expectedName)
    {
        bool resolved = ProvinceCatalog.TryResolve(value, out Province province);

        Assert.True(resolved);
        Assert.Equal(expectedName, province.Name);
    }

    [Theory]
    [InlineData("云公司")]
    [InlineData("香港")]
    [InlineData("")]
    public void TryResolveRejectsUnsupportedValues(string value)
    {
        bool resolved = ProvinceCatalog.TryResolve(value, out _);

        Assert.False(resolved);
    }
}

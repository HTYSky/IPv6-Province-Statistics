using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Templates;

public sealed class TemplateSanitizerTests
{
    [Fact]
    public void Sanitize_clears_all_values_and_passes_validation()
    {
        var source = TemplateFixtures.CreatePopulatedTemplate();
        var output = TempFiles.Next("sanitized.xlsx");
        TemplateSanitizer.Sanitize(source, output);
        Assert.Empty(TemplateValidator.Validate(output));
    }
}

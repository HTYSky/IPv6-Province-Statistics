using System.Reflection;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

public sealed class TemplateResourceProvider
{
    private readonly Assembly _assembly = typeof(TemplateResourceProvider).Assembly;

    public byte[] ReadTemplate()
    {
        using var stream = _assembly.GetManifestResourceStream("Ipv6ProvinceStatistics.Template.xlsx")
            ?? throw new InvalidDataException("Embedded template missing.");
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    public string ReadTemplateSha256()
    {
        using var reader = new StreamReader(
            _assembly.GetManifestResourceStream("Ipv6ProvinceStatistics.Template.sha256")
                ?? throw new InvalidDataException("Embedded template sha256 missing."));
        return reader.ReadToEnd().Trim();
    }
}

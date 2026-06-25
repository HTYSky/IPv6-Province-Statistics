using System.Security.Cryptography;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

if (args.Length != 4 || args[0] != "sanitize")
{
    Console.Error.WriteLine("Usage: sanitize <input.xlsx> <output.xlsx> <output.sha256>");
    return 2;
}

Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(args[2]))!);
TemplateSanitizer.Sanitize(args[1], args[2]);
using var stream = File.OpenRead(args[2]);
await File.WriteAllTextAsync(args[3], Convert.ToHexString(SHA256.HashData(stream)) + Environment.NewLine);
return 0;

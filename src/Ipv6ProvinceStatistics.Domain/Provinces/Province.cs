namespace Ipv6ProvinceStatistics.Domain.Provinces;

public readonly record struct Province(string Name)
{
    public override string ToString() => Name;
}

namespace Ipv6ProvinceStatistics.Application.Models;
public enum ProcessingStage { Snapshotting, Identifying, Validating, Generating, Verifying, Publishing }
public sealed record ProcessingProgress(ProcessingStage Stage, int Completed, int Total, string Message);

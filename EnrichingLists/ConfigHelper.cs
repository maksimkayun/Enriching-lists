using Microsoft.Extensions.Configuration;

namespace Enriching_lists;

public static class ConfigHelper
{
    public static IConfigurationRoot ConfigurationRoot =
        new ConfigurationBuilder()
            .AddJsonFile("appsettings.json")
            .AddEnvironmentVariables()
            .Build();

    public static Dictionary<string, int> Headers
        => ConfigurationRoot
            .GetSection("Headers")
            .Get<string[]>()
            .ToDictionary(
                k => k.Split(" == ").First(),
                v => int.Parse(v.Split(" == ").Last())
            );
}
using System;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using HttpListenerService;
using Microsoft.Extensions.Configuration;

class Program
{
    static async Task Main(string[] args)
    {
        var config = LoadConfiguration();
        HttpListenerClass httpListenerClass = new HttpListenerClass(config);
        await httpListenerClass.Start();

    }

    static IConfigurationRoot LoadConfiguration()
    {
        var configBuilder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

        return configBuilder.Build();
    }
}

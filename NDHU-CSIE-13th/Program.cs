using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using NDHU_CSIE_13th.Repos;
using NDHU_CSIE_13th.Services;
using NLog.Web;

internal class Program
{
    private static void Main(string[] args)
    {
        var host = Host.CreateDefaultBuilder(args)
              .ConfigureAppConfiguration((context, config) =>
              {
                  config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
              })
              .ConfigureServices((context, services) => 
              {
                  services.AddScoped<ExcelRepo>();
                  services.AddScoped<SettlementService>();
              })
              .UseNLog()
              .Build();

        var service = host.Services.GetRequiredService<SettlementService>();

        service.Start();
    }
}
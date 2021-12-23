using System;
using System.IO;
using System.Threading.Tasks;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;


using VersionControlHelper.Config;
using VersionControlHelper.Domain;
using VersionControlHelper.Service;

namespace VersionControlHelper
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var builder = new HostBuilder()
               .ConfigureAppConfiguration((hostingContext, config) =>
               {
                   config.SetBasePath(Path.Combine(AppContext.BaseDirectory));
                   config.AddJsonFile("appsettings.json", optional: false, true);
                   config.AddEnvironmentVariables();

                   if (args != null)
                   {
                       config.AddCommandLine(args);
                   }
               })
               .ConfigureServices((hostContext, services) =>
               {
                   services.AddOptions();

                   services.Configure<AppConfig>(hostContext.Configuration.GetSection("AppConfig"));

                   services.AddScoped<ITfsCommitHistoryService, TfsCommitHistoryService>();
                   services.AddScoped<ICommitHistoryService, TfsCommitHistoryService>();

                   services.AddScoped<IGitCommitHistoryService, GitCommitHistoryService>();
                   services.AddScoped<ICommitHistoryService, GitCommitHistoryService>();

                   services.AddHostedService<App>();
               })
               .ConfigureLogging((hostingContext, logging) =>
               {
                   logging.AddConfiguration(hostingContext.Configuration.GetSection("Logging"));
                   logging.AddConsole();
               });

            await builder.RunConsoleAsync();
        }
    }
}

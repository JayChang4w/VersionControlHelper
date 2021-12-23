using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.VisualStudio.Services.Common;

using VersionControlHelper.Config;
using VersionControlHelper.Domain;

namespace VersionControlHelper
{
    public class App : IHostedService
    {
        private readonly ILogger<App> _logger;
        private readonly IHostApplicationLifetime _appLifetime;

        private readonly AppConfig _config;
        private readonly ICommitHistoryService _historyService;

        public App(ILogger<App> logger,
                   IHostApplicationLifetime appLifetime,
                   IOptionsSnapshot<AppConfig> configSnapshot,
                   IEnumerable<ICommitHistoryService> historyServices)
        {
            this._logger = logger;
            this._appLifetime = appLifetime;

            this._config = configSnapshot?.Value ?? throw new ConfigFileException(nameof(AppConfig));
            this._historyService = historyServices.Where(w => w.Type == this._config.Type).LastOrDefault() ?? throw new NotImplementedException();
        }

        public async Task StartAsync(CancellationToken cancellationToken)
        {
            //this._logger.LogInformation($"App running at: {DateTimeOffset.Now}");

            await Task.Yield();

            Console.WriteLine($"專案: {this._config.Project}");

            Console.WriteLine($"版控類型: {this._config.Type}");

            Console.WriteLine($"正在匯出資料至「{System.IO.Path.Combine(AppContext.BaseDirectory, this._config.ExportPath)}」...");

            await this._historyService.ExportExcelAsync();

            Console.WriteLine("匯出完成");

            Console.WriteLine("press any key to exit...");

            Console.ReadKey();

            this._appLifetime.StopApplication();
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            this._logger.LogWarning($"App stopped at: {DateTimeOffset.Now}");

            return Task.CompletedTask;
        }
    }
}

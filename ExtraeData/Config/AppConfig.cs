using Microsoft.Extensions.Configuration;
using System;

namespace ExtraeData.Config
{
    public sealed class AppConfig
    {
        public string ConnectionString { get; }
        public string Username { get; }
        public string Password { get; }
        public int MaxRetries { get; }
        public int BanMinutes { get; }
        public int CooldownSeconds { get; }

        private AppConfig(string cs, string user, string pass, int maxRetries, int banMinutes, int cooldownSeconds)
        {
            ConnectionString = cs;
            Username = user;
            Password = pass;
            MaxRetries = maxRetries;
            BanMinutes = banMinutes;
            CooldownSeconds = cooldownSeconds;
        }

        public static AppConfig Load()
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false)
                .Build();

            var cs = config.GetConnectionString("Server25")
                ?? throw new Exception("Falta ConnectionStrings:Server25 en appsettings.json");

            var user = config["Veritrade:Username"]
                ?? throw new Exception("Falta Veritrade:Username en appsettings.json");

            var pass = config["Veritrade:Password"]
                ?? throw new Exception("Falta Veritrade:Password en appsettings.json");

            int maxRetries = int.TryParse(config["Rpa:MaxRetries"], out var mr) ? mr : 3;
            int banMinutes = int.TryParse(config["Rpa:BanMinutes"], out var bm) ? bm : 20;
            int cooldownSeconds = int.TryParse(config["Rpa:CooldownSeconds"], out var csd) ? csd : 15;

            return new AppConfig(cs, user, pass, maxRetries, banMinutes, cooldownSeconds);
        }
    }
}
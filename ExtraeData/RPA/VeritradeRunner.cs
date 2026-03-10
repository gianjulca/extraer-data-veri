/*using ExtraeData.Config;
using ExtraeData.Constants;
using ExtraeData.Data;
using ExtraeData.Services;
using ExtraeData.Web;
using ExtraeData.Workflows;
using System;
using System.IO;
using System.Threading.Tasks;

namespace ExtraeData.Rpa
{
    public static class VeritradeRunner
    {
        public static async Task RunAsync()
        {
            var cfg = AppConfig.Load();

            var downloadsDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Descargas"
            );
            Directory.CreateDirectory(downloadsDir);

            var reader = new VeritradeExcelReader();
            var repo = new VeritradeSqlRepository(cfg.ConnectionString);

            await using var session = new VeritradeWebSession();
            await session.StartAsync();
            await session.LoginAsync(cfg.Username, cfg.Password);

            // 1) Importaciones Perú
            await ImportacionesPeruWorkflow.RunAsync(session.Page, downloadsDir, reader, repo);

            // 2) Importaciones otros países
            await ImportacionesOtrosPaisesWorkflow.RunAsync(session.Page, downloadsDir, reader, repo);

            // 3) Exportaciones Perú
            await ExportacionesPeruWorkflow.RunAsync(session.Page, downloadsDir, reader, repo);

            await session.LogoutAsync();
        }
    }
}*/

using ExtraeData.Config;
using ExtraeData.Data;
using ExtraeData.Services;
using ExtraeData.Web;
using ExtraeData.Workflows;
using System;
using System.IO;
using System.Threading.Tasks;

namespace ExtraeData.Rpa
{
    public static class VeritradeRunner
    {
        public static async Task RunAsync()
        {
            var cfg = AppConfig.Load();

            var downloadsDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Descargas"
            );
            Directory.CreateDirectory(downloadsDir);

            var reader = new VeritradeExcelReader();
            var repo = new VeritradeSqlRepository(cfg.ConnectionString);

            for (int attempt = 1; attempt <= cfg.MaxRetries; attempt++)
            {
                await using var session = new VeritradeWebSession();
                bool completed = false;

                try
                {
                    await session.StartAsync();
                    await session.LoginAsync(cfg.Username, cfg.Password);

                    await session.EnsureHealthyOrThrowAsync("after-login");

                    await ImportacionesPeruWorkflow.RunAsync(session.Page, downloadsDir, reader, repo);
                    await session.EnsureHealthyOrThrowAsync("after-imp-peru");

                    await ImportacionesOtrosPaisesWorkflow.RunAsync(session.Page, downloadsDir, reader, repo);
                    await session.EnsureHealthyOrThrowAsync("after-imp-otros");

                    await ExportacionesPeruWorkflow.RunAsync(session.Page, downloadsDir, reader, repo);
                    await session.EnsureHealthyOrThrowAsync("after-exp-peru");

                    completed = true;
                }

                catch (Exception ex)
                {
                    Console.WriteLine($"[RUN-FAIL] attempt={attempt}/{cfg.MaxRetries} => {ex.Message}");

                    try
                    {
                        Console.WriteLine("[RUNNER] Intentando cerrar sesión web tras fallo...");
                        await session.LogoutAsync();
                        Console.WriteLine("[RUNNER] Logout web OK tras fallo.");
                    }
                    catch (Exception logoutEx)
                    {
                        Console.WriteLine($"[RUNNER] No se pudo cerrar sesión web tras fallo: {logoutEx.Message}");
                    }

                    var msg = ex.ToString();

                    var isInputIssue = msg.Contains("No se pudo mantener el valor", StringComparison.OrdinalIgnoreCase);
                    var isSingleSession = msg.Contains("VERITRADE_SINGLE_SESSION", StringComparison.OrdinalIgnoreCase);
                    var isBrokenSession =
                        msg.Contains("VERITRADE_SERVER_ERROR", StringComparison.OrdinalIgnoreCase) ||
                        msg.Contains("502", StringComparison.OrdinalIgnoreCase) ||
                        msg.Contains("Server Error", StringComparison.OrdinalIgnoreCase) ||
                        msg.Contains("The wait operation timed out", StringComparison.OrdinalIgnoreCase) ||
                        msg.Contains("404", StringComparison.OrdinalIgnoreCase) ||
                        msg.Contains("Not Found", StringComparison.OrdinalIgnoreCase);

                    var isTimeout = msg.Contains("Timeout", StringComparison.OrdinalIgnoreCase);

                    if (isInputIssue)
                    {
                        Console.WriteLine("[NO-RETRY] Error funcional de input. No se reintenta sesión.");
                        throw;
                    }

                    if (isSingleSession || isBrokenSession)
                    {
                        Console.WriteLine($"[BAN-WAIT] sesión rota o bloqueada. Esperando {cfg.BanMinutes} min...");
                        await Task.Delay(TimeSpan.FromMinutes(cfg.BanMinutes));
                    }
                    else if (isTimeout)
                    {
                        Console.WriteLine($"[COOLDOWN] esperando {cfg.CooldownSeconds}s...");
                        await Task.Delay(TimeSpan.FromSeconds(cfg.CooldownSeconds));
                    }

                    if (attempt == cfg.MaxRetries)
                        throw;
                }
                finally
                {
                    if (completed)
                    {
                        try
                        {
                            await session.LogoutAsync();
                            Console.WriteLine("Sesión cerrada correctamente.");
                        }
                        catch (Exception logoutEx)
                        {
                            Console.WriteLine($"[LOGOUT-WARN] El proceso terminó OK, pero falló el cierre de sesión: {logoutEx.Message}");
                        }
                    }
                }

                if (completed)
                    return;
            }
        }
    }
}

/*using Microsoft.Playwright;
using System;
using System.Threading.Tasks;

namespace ExtraeData.Web
{
    public sealed class VeritradeWebSession : IAsyncDisposable
    {
        public IPlaywright Playwright { get; private set; } = default!;
        public IBrowser Browser { get; private set; } = default!;
        public IBrowserContext Context { get; private set; } = default!;
        public IPage Page { get; private set; } = default!;

        public async Task StartAsync()
        {
            Playwright = await Microsoft.Playwright.Playwright.CreateAsync();

            Browser = await Playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
            {
                Headless = false,
                SlowMo = 90,
                Channel = "chrome",
                Args = new[]
                {
                    "--window-position=0,0",
                    "--window-size=1520,730",
                    "--disable-dev-shm-usage",
                    "--disable-features=TranslateUI"
                }
            });

            Context = await Browser.NewContextAsync(new BrowserNewContextOptions
            {
                Locale = "es-PE", 
                AcceptDownloads = true,
                ViewportSize = new ViewportSize { Width = 1520, Height = 730 },
                UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            });

            await Context.AddInitScriptAsync("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});");

            Page = await Context.NewPageAsync();
        }

        public async Task LoginAsync(string user, string pass)
        {
            await Page.GotoAsync("https://www.veritradecorp.com/",
                new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded, Timeout = 120000 });

            await Page.ClickAsync("#login");
            await Page.WaitForSelectorAsync("#txtCodUsuario", new() { Timeout = 60000 });

            await Page.FillAsync("#txtCodUsuario", user);
            await Page.FillAsync("#txtPassword", pass);

            var terms = Page.Locator("input[type='checkbox']");
            if (await terms.CountAsync() > 0)
            {
                var first = terms.First;
                if (!await first.IsCheckedAsync())
                    await first.CheckAsync();
            }

            await Page.ClickAsync("text=Login >");

            await Page.WaitForURLAsync("/es/mis-busquedas**", new() { Timeout = 120000 });
            await Page.WaitForSelectorAsync("#btnBuscar", new() { Timeout = 120000 });
        }

        public async Task LogoutAsync()
        {
            await Page.EvaluateAsync("window.scrollTo(0, 0)");

            var gear = Page.Locator("span.glyphicon.glyphicon-cog").First;
            await gear.WaitForAsync(new() { Timeout = 60000 });
            await gear.ClickAsync(new() { Force = true });

            await Page.GetByRole(AriaRole.Link, new() { Name = "Cerrar Sesión" })
                .ClickAsync(new() { Force = true });

            await Page.WaitForSelectorAsync("#login", new() { Timeout = 120000 });
        }

        public async ValueTask DisposeAsync()
        {
            try { if (Browser != null) await Browser.CloseAsync(); } catch { }
            try { Playwright?.Dispose(); } catch { }
        }
    }
}*/

using Microsoft.Playwright;
using System;
using System.Threading.Tasks;

namespace ExtraeData.Web
{
    public sealed class VeritradeWebSession : IAsyncDisposable
    {
        public IPlaywright Playwright { get; private set; } = default!;
        public IBrowser Browser { get; private set; } = default!;
        public IBrowserContext Context { get; private set; } = default!;
        public IPage Page { get; private set; } = default!;
        public async Task StartAsync()
        {
            Playwright = await Microsoft.Playwright.Playwright.CreateAsync();

            Browser = await Playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
            {
                Headless = false,
                SlowMo = 90,
                Channel = "chrome",
                Args = new[]
                {
                    "--window-position=0,0",
                    "--window-size=1520,730",
                    "--disable-dev-shm-usage",
                    "--disable-features=TranslateUI"
                }
            });

            Context = await Browser.NewContextAsync(new BrowserNewContextOptions
            {
                Locale = "es-PE",
                AcceptDownloads = true,
                ViewportSize = new ViewportSize { Width = 1520, Height = 730 },
                UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            });

            await Context.AddInitScriptAsync("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});");

            Page = await Context.NewPageAsync();
        }
        public async Task LogoutAsync()
        {
            if (Page == null) return;

            Console.WriteLine("[LOGOUT] Iniciando cierre de sesión...");

            await Page.EvaluateAsync("window.scrollTo(0, 0)");
            await Page.WaitForTimeoutAsync(500);

            var gearLink = Page.Locator("a:has(span.glyphicon.glyphicon-cog)").First;
            if (await gearLink.CountAsync() == 0)
                gearLink = Page.Locator("span.glyphicon.glyphicon-cog").Locator("xpath=ancestor::a[1]").First;

            await gearLink.WaitForAsync(new() { Timeout = 15000 });
            Console.WriteLine("[LOGOUT] Engranaje encontrado");

            await gearLink.ClickAsync();
            Console.WriteLine("[LOGOUT] Click en engranaje OK");

            var logoutLink = Page.Locator("a:has-text('Cerrar Sesión')").First;
            await logoutLink.WaitForAsync(new() { Timeout = 15000 });
            Console.WriteLine("[LOGOUT] Opción 'Cerrar Sesión' encontrada");

            await logoutLink.ClickAsync();
            Console.WriteLine("[LOGOUT] Click en 'Cerrar Sesión' OK");

            await Page.WaitForSelectorAsync("#login", new() { Timeout = 30000 });
            Console.WriteLine("[LOGOUT] Sesión cerrada correctamente");
        }
        public async Task LoginAsync(string user, string pass)
        {
            await Page.GotoAsync("https://www.veritradecorp.com/",
                new PageGotoOptions
                {
                    WaitUntil = WaitUntilState.DOMContentLoaded,
                    Timeout = 120000
                });

            await Page.WaitForTimeoutAsync(1500);

            if (await IsServerErrorPageAsync())
                throw new Exception("VERITRADE_SERVER_ERROR: error de servidor antes del login.");

            if (await IsSingleSessionPageAsync())
            {
                await ClickVolverAVeritradeIfExistsAsync();
                await Page.WaitForTimeoutAsync(1500);

                if (await IsSingleSessionPageAsync())
                    throw new Exception("VERITRADE_SINGLE_SESSION: usuario en línea, debe esperar 20 minutos.");
            }

            await Page.ClickAsync("#login");
            await Page.WaitForSelectorAsync("#txtCodUsuario", new() { Timeout = 60000 });

            await Page.FillAsync("#txtCodUsuario", user);
            await Page.FillAsync("#txtPassword", pass);

            var terms = Page.Locator("input[type='checkbox']");
            if (await terms.CountAsync() > 0)
            {
                var first = terms.First;
                if (!await first.IsCheckedAsync())
                    await first.CheckAsync();
            }

            await Page.ClickAsync("text=Login >");

            await Page.WaitForTimeoutAsync(2500);

            if (await IsSingleSessionPageAsync())
                throw new Exception("VERITRADE_SINGLE_SESSION: usuario en línea luego del login.");

            if (await IsServerErrorPageAsync())
                throw new Exception("VERITRADE_SERVER_ERROR: error de servidor luego del login.");

            await Page.WaitForURLAsync("**/es/mis-busquedas**", new() { Timeout = 120000 });
            await Page.WaitForSelectorAsync("#btnBuscar", new() { Timeout = 120000 });
        }
        public async ValueTask DisposeAsync()
        {
            try { if (Page != null) await Page.CloseAsync(); } catch { }
            try { if (Context != null) await Context.CloseAsync(); } catch { }
            try { if (Browser != null) await Browser.CloseAsync(); } catch { }
            try { Playwright?.Dispose(); } catch { }
        }
        public async Task<bool> IsServerErrorPageAsync()
        {
            try
            {
                var url = Page.Url ?? "";
                var title = (await Page.TitleAsync()) ?? "";

                var body = Page.Locator("body");
                var bodyText = "";
                if (await body.CountAsync() > 0)
                    bodyText = (await body.InnerTextAsync()) ?? "";

                bool is502 = title.Contains("502", StringComparison.OrdinalIgnoreCase) ||
                             bodyText.Contains("502 Bad Gateway", StringComparison.OrdinalIgnoreCase);

                bool isServerError = bodyText.Contains("Server Error in '/' Application", StringComparison.OrdinalIgnoreCase) ||
                                     bodyText.Contains("The wait operation timed out", StringComparison.OrdinalIgnoreCase);

                bool is404 = title.Contains("404", StringComparison.OrdinalIgnoreCase) ||
                             bodyText.Contains("HTTP Error 404.0 - Not Found", StringComparison.OrdinalIgnoreCase) ||
                             bodyText.Contains("Not Found", StringComparison.OrdinalIgnoreCase);

                bool looksBroken = is502 || isServerError || is404;

                if (!looksBroken &&
                    url.Contains("DownloadFile", StringComparison.OrdinalIgnoreCase) &&
                    (title.Contains("Bad Gateway", StringComparison.OrdinalIgnoreCase) ||
                     bodyText.Contains("Bad Gateway", StringComparison.OrdinalIgnoreCase)))
                {
                    looksBroken = true;
                }

                return looksBroken;
            }
            catch
            {
                return true;
            }
        }
        public async Task EnsureHealthyOrThrowAsync(string where)
        {
            if (await IsServerErrorPageAsync())
                throw new Exception($"VERITRADE_SERVER_ERROR at {where}: se detectó pantalla 502/ServerError.");
        }
        private static bool ContainsAny(string s, params string[] keys)
        {
            s = (s ?? "").ToLowerInvariant();
            return keys.Any(k => s.Contains((k ?? "").ToLowerInvariant()));
        }
        private async Task<bool> IsSingleSessionPageAsync()
        {
            try
            {
                var url = Page.Url ?? "";

                string bodyText = "";
                var body = Page.Locator("body");
                if (await body.CountAsync() > 0)
                    bodyText = (await body.InnerTextAsync()) ?? "";

                return ContainsAny(bodyText,
                    "ya existe el usuario en línea",
                    "solo permite una sesión abierta",
                    "deberá esperar 20 minutos",
                    "volver a veritrade"
                ) || ContainsAny(url, "/error/mostrar");
            }
            catch (PlaywrightException)
            {
                return false;
            }
            catch
            {
                return false;
            }
        }
        private async Task ClickVolverAVeritradeIfExistsAsync()
        {
            try
            {
                var btn = Page.Locator("a.btn.btnVeritrade:visible, text=VOLVER A VERITRADE").First;
                if (await btn.CountAsync() > 0)
                    await btn.ClickAsync(new() { Force = true });
            }
            catch
            {
                // no romper si no existe o la página cambia
            }
        }

    }
}



using Microsoft.Playwright;
using System;
using System.Threading.Tasks;

class Program
{
    static async Task Main()
    {
        using var playwright = await Playwright.CreateAsync();

        // Usa Chrome real (no Chromium) si lo tienes instalado
        var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
        {
            Headless = false,
            SlowMo = 50,
            Channel = "chrome",
            Args = new[]
            {
                "--disable-blink-features=AutomationControlled",
                "--start-maximized"
            }
        });

        var context = await browser.NewContextAsync(new BrowserNewContextOptions
        {
            Locale = "es-PE",
            ViewportSize = null, // para que use tamaño real de ventana
            UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        });

        // Quita "webdriver" (algunos sitios lo detectan)
        await context.AddInitScriptAsync("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});");

        var page = await context.NewPageAsync();

        // Log de requests y responses (solo lo más útil)
        page.Response += async (_, resp) =>
        {
            if (resp.Url.Contains("/ajax/auth", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"AUTH RESP => {resp.Status} {resp.Url}");
                try
                {
                    var body = await resp.TextAsync();
                    Console.WriteLine($"AUTH BODY => {body}");
                }
                catch { /* a veces no deja leer */ }
            }
        };

        await page.GotoAsync("https://www.veritradecorp.com/",
            new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded, Timeout = 120000 });

        await page.ClickAsync("#login");

        await page.WaitForSelectorAsync("#txtCodUsuario", new() { Timeout = 60000 });

        // OJO: acá pon tus valores (ideal: variables de entorno)
        await page.FillAsync("#txtCodUsuario", "correo");
        await page.FillAsync("#txtPassword", "contraseña");

        // Si existe checkbox de términos, márcalo explícitamente
        var terms = page.Locator("input[type='checkbox']");
        if (await terms.CountAsync() > 0)
        {
            // intenta marcar el primero si no está marcado
            var first = terms.First;
            if (!await first.IsCheckedAsync())
                await first.CheckAsync();
        }

        // Click login y espera específicamente la respuesta /ajax/auth
        await page.ClickAsync("text=Login >");

        await page.WaitForTimeoutAsync(300000);
        Console.WriteLine("Listo: revisa consola para AUTH RESP/BODY.");
    }
}

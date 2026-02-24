/*using Microsoft.Playwright;
using System;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Linq;

class Program
{
    static async Task Main()
    {
        string? excelPath = null;

        // ✅ Credenciales hardcodeadas (pruebas)
        var user = "";
        var pass = "";

        // ✅ Guardar descarga en "Descargas" (DEFINIR SOLO UNA VEZ)
        var downloadsDir = Path.Combine(
       Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
       "Descargar"
           );
        Directory.CreateDirectory(downloadsDir);

        using var playwright = await Playwright.CreateAsync();

        var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
        {
            Headless = false,
            SlowMo = 50,
            Channel = "chrome",
            Args = new[] { "--start-maximized" }
        });

        var context = await browser.NewContextAsync(new BrowserNewContextOptions
        {
            Locale = "es-PE",
            ViewportSize = null,
            UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            AcceptDownloads = true
        });

        var page = await context.NewPageAsync();

        await page.GotoAsync("https://www.veritradecorp.com/",
            new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded, Timeout = 120000 });

        // 1) LOGIN
        await page.ClickAsync("#login");
        await page.WaitForSelectorAsync("#txtCodUsuario", new() { Timeout = 60000 });

        await page.FillAsync("#txtCodUsuario", user);
        await page.FillAsync("#txtPassword", pass);

        var terms = page.Locator("input[type='checkbox']");
        if (await terms.CountAsync() > 0)
        {
            var first = terms.First;
            if (!await first.IsCheckedAsync())
                await first.CheckAsync();
        }

        await page.ClickAsync("text=Login >");

        await page.GetByRole(AriaRole.Link, new() { Name = "Mis Productos" })
                  .WaitForAsync(new() { Timeout = 120000 });

        // 2) Click botón azul "Mis Productos" (abre modal)
        await page.GetByRole(AriaRole.Button, new() { Name = "Mis Productos" }).ClickAsync();

        // 3) Modal "Mis Productos" (más específico)
        var modal = page.Locator(".modal-dialog:has-text('Mis Productos'), .modal-content:has-text('Mis Productos'), .modal:has-text('Mis Productos')").First;
        await modal.WaitForAsync(new() { Timeout = 60000 });

        // 3.1) Tomar un checkbox visible dentro del modal
        var cbxs = modal.Locator("input[type='checkbox']:visible");
        await cbxs.First.WaitForAsync(new() { Timeout = 60000 });
        var cbx = cbxs.First;

        try
        {
            await cbx.CheckAsync(new() { Force = true });
        }
        catch
        {
            await cbx.ClickAsync(new() { Force = true });
        }

        // 3.2) Click "Agregar a Filtros" dentro del modal
        await modal.GetByRole(AriaRole.Button, new() { Name = "Agregar a Filtros" })
                  .ClickAsync(new() { Force = true });

        // 3.3) Esperar que el modal se cierre
        await modal.WaitForAsync(new() { State = WaitForSelectorState.Hidden, Timeout = 60000 });

        // 4) Buscar
        await page.GetByRole(AriaRole.Button, new() { Name = "BUSCAR" }).ClickAsync();

        await page.GetByRole(AriaRole.Tab, new() { Name = "Resumen" })
                  .WaitForAsync(new() { Timeout = 120000 });

        await page.GetByRole(AriaRole.Tab, new() { Name = "Resumen" }).ClickAsync();

        // 5) Click en "Ver Registros" (el número debajo de REGISTROS)
        var verRegistros = page.Locator("#tbodyResumenPartida a.lnkVerRegistros").First;
        await verRegistros.WaitForAsync(new() { Timeout = 60000 });
        await verRegistros.ClickAsync(new() { Force = true });

        // 5.1) Esperar el modal de detalle (Partida...)
        var detailModal = page.Locator(".modal-dialog:has-text('Partida'), .modal-content:has-text('Partida')").First;
        await detailModal.WaitForAsync(new() { Timeout = 60000 });

        // 5.2) Esperar que exista el botón Excel dentro del modal
        var excelAnchor = detailModal.Locator("#downloadFileVerRegistro").First;
        await excelAnchor.WaitForAsync(new() { Timeout = 120000 });

        // 6) Descargar Excel (puede demorar varios minutos)
        var downloadTask = page.WaitForDownloadAsync(new() { Timeout = 300000 }); // 5 min

        await excelAnchor.ClickAsync(new() { Force = true });

        // Espera el download y guarda en "Descargar"
        var download = await downloadTask;

        var filename = download.SuggestedFilename;
        if (string.IsNullOrWhiteSpace(filename))
            filename = "veritrade.xlsx";

        var fullPath = Path.Combine(downloadsDir, $"{DateTime.Now:yyyyMMdd_HHmmss}_{filename}");
        await download.SaveAsAsync(fullPath);

        Console.WriteLine($"✅ Excel descargado y guardado en: {fullPath}");

        excelPath = fullPath;


        // 7) Cerrar el modal con la "X"
        var closeBtn = detailModal.Locator("button.close, .modal-header button, .close").First;
        await closeBtn.ClickAsync(new() { Force = true });

        // Espera que el modal desaparezca
        await detailModal.WaitForAsync(new() { State = WaitForSelectorState.Hidden, Timeout = 60000 });

        // 8) Scroll arriba para llegar al engranaje
        await page.EvaluateAsync("window.scrollTo(0, 0)");

        // 9) Abrir menú del engranaje (según tu HTML: span.glyphicon.glyphicon-cog)
        var gear = page.Locator("span.glyphicon.glyphicon-cog").First;
        await gear.WaitForAsync(new() { Timeout = 60000 });
        await gear.ClickAsync(new() { Force = true });

        // 10) Click en "Cerrar Sesión"
        await page.GetByRole(AriaRole.Link, new() { Name = "Cerrar Sesión" })
                  .ClickAsync(new() { Force = true });

        // (Opcional) Esperar que vuelva al home/login
        await page.WaitForSelectorAsync("#login", new() { Timeout = 120000 });

        Console.WriteLine("✅ Sesión cerrada correctamente.");


        //leer excel y mostrar datos

        if (!string.IsNullOrWhiteSpace(excelPath) && File.Exists(excelPath))
        {
            PrintExcelColumns(excelPath, "Partida Aduanera", "Importador");
        }
        else
        {
            Console.WriteLine("⚠️ No se encontró el archivo Excel para leer.");
        }

    }

    static void PrintExcelColumns(string filePath, string col1Name, string col2Name)
    {
        using var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheets.First();

        // Encuentra la fila de encabezados buscando col1Name
        var headerRow = ws.RowsUsed()
            .FirstOrDefault(r => r.CellsUsed().Any(c =>
                string.Equals(c.GetString().Trim(), col1Name, StringComparison.OrdinalIgnoreCase)));

        if (headerRow == null)
            throw new Exception($"No se encontró la fila de encabezados con '{col1Name}'.");

        int headerRowNum = headerRow.RowNumber();

        int col1 = headerRow.CellsUsed()
            .First(c => string.Equals(c.GetString().Trim(), col1Name, StringComparison.OrdinalIgnoreCase))
            .Address.ColumnNumber;

        int col2 = headerRow.CellsUsed()
            .First(c => string.Equals(c.GetString().Trim(), col2Name, StringComparison.OrdinalIgnoreCase))
            .Address.ColumnNumber;

        var lastRow = ws.LastRowUsed().RowNumber();

        Console.WriteLine("---------------------------------------------------");
        Console.WriteLine($"📄 Leyendo Excel: {filePath}");
        Console.WriteLine($"Headers fila {headerRowNum} | {col1Name}=col {col1} | {col2Name}=col {col2}");
        Console.WriteLine("---------------------------------------------------");

        for (int r = headerRowNum + 1; r <= lastRow; r++)
        {
            var v1 = ws.Cell(r, col1).GetString().Trim();
            var v2 = ws.Cell(r, col2).GetString().Trim();

            if (string.IsNullOrWhiteSpace(v1) && string.IsNullOrWhiteSpace(v2))
                continue;

            Console.WriteLine($"{col1Name}: {v1} | {col2Name}: {v2}");
        }
    }

} */

using Microsoft.Playwright;
using System;
using System.IO;
using System.Threading.Tasks;
using ExtraeData.Services;
using ExtraeData.Data;
using Microsoft.Extensions.Configuration;

class Program
{
    static async Task Main()
    {
        //conexion a SQL
        var config = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false)
    .Build();

        var cs = config.GetConnectionString("Server25")
                 ?? throw new Exception("Falta ConnectionStrings:Server25 en appsettings.json");
        //finSQL

        string? excelPath = null;

        // Credenciales hardcodeadas (pruebas)
        var user = "rossana.iglesias@nogasa.com.pe";
        var pass = "Vistony2020";

        // Carpeta donde se guardará el Excel (tu carpeta: "Descargar")
        var downloadsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Descargar"
        );
        Directory.CreateDirectory(downloadsDir);

        using var playwright = await Playwright.CreateAsync();

        var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
        {
            Headless = false,
            SlowMo = 50,
            Channel = "chrome",
            Args = new[] { "--start-maximized" }
        });

        var context = await browser.NewContextAsync(new BrowserNewContextOptions
        {
            Locale = "es-PE",
            ViewportSize = null,
            UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            AcceptDownloads = true
        });

        var page = await context.NewPageAsync();

        await page.GotoAsync("https://www.veritradecorp.com/",
            new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded, Timeout = 120000 });

        // 1) LOGIN
        await page.ClickAsync("#login");
        await page.WaitForSelectorAsync("#txtCodUsuario", new() { Timeout = 60000 });

        await page.FillAsync("#txtCodUsuario", user);
        await page.FillAsync("#txtPassword", pass);

        var terms = page.Locator("input[type='checkbox']");
        if (await terms.CountAsync() > 0)
        {
            var first = terms.First;
            if (!await first.IsCheckedAsync())
                await first.CheckAsync();
        }

        await page.ClickAsync("text=Login >");

        await page.GetByRole(AriaRole.Link, new() { Name = "Mis Productos" })
                  .WaitForAsync(new() { Timeout = 120000 });

        // 2) Click botón azul "Mis Productos" (abre modal)
        await page.GetByRole(AriaRole.Button, new() { Name = "Mis Productos" }).ClickAsync();

        // 3) Modal "Mis Productos"
        var modal = page.Locator(".modal-dialog:has-text('Mis Productos'), .modal-content:has-text('Mis Productos'), .modal:has-text('Mis Productos')").First;
        await modal.WaitForAsync(new() { Timeout = 60000 });

        // 3.1) Marcar un checkbox visible (prueba)
        var cbxs = modal.Locator("input[type='checkbox']:visible");
        await cbxs.First.WaitForAsync(new() { Timeout = 60000 });
        var cbx = cbxs.First;

        try { await cbx.CheckAsync(new() { Force = true }); }
        catch { await cbx.ClickAsync(new() { Force = true }); }

        // 3.2) "Agregar a Filtros"
        await modal.GetByRole(AriaRole.Button, new() { Name = "Agregar a Filtros" })
                  .ClickAsync(new() { Force = true });

        // 3.3) Esperar cierre del modal
        await modal.WaitForAsync(new() { State = WaitForSelectorState.Hidden, Timeout = 60000 });

        // 4) Buscar
        await page.GetByRole(AriaRole.Button, new() { Name = "BUSCAR" }).ClickAsync();

        await page.GetByRole(AriaRole.Tab, new() { Name = "Resumen" })
                  .WaitForAsync(new() { Timeout = 120000 });

        await page.GetByRole(AriaRole.Tab, new() { Name = "Resumen" }).ClickAsync();

        // 5) Click "Ver Registros"
        var verRegistros = page.Locator("#tbodyResumenPartida a.lnkVerRegistros").First;
        await verRegistros.WaitForAsync(new() { Timeout = 60000 });
        await verRegistros.ClickAsync(new() { Force = true });

        // 5.1) Modal de detalle (Partida...)
        var detailModal = page.Locator(".modal-dialog:has-text('Partida'), .modal-content:has-text('Partida')").First;
        await detailModal.WaitForAsync(new() { Timeout = 60000 });

        // 5.2) Botón Excel
        var excelAnchor = detailModal.Locator("#downloadFileVerRegistro").First;
        await excelAnchor.WaitForAsync(new() { Timeout = 120000 });

        // 6) Descargar Excel (puede demorar)
        var downloadTask = page.WaitForDownloadAsync(new() { Timeout = 300000 }); // 5 min
        await excelAnchor.ClickAsync(new() { Force = true });

        var download = await downloadTask;
        var filename = download.SuggestedFilename;
        if (string.IsNullOrWhiteSpace(filename))
            filename = "veritrade.xlsx";

        var fullPath = Path.Combine(downloadsDir, $"{DateTime.Now:yyyyMMdd_HHmmss}_{filename}");
        await download.SaveAsAsync(fullPath);

        Console.WriteLine($"Excel descargado y guardado en: {fullPath}");
        excelPath = fullPath;

        // 7) Cerrar modal detalle
        var closeBtn = detailModal.Locator("button.close, .modal-header button, .close").First;
        await closeBtn.ClickAsync(new() { Force = true });
        await detailModal.WaitForAsync(new() { State = WaitForSelectorState.Hidden, Timeout = 60000 });

        // 8) Scroll arriba
        await page.EvaluateAsync("window.scrollTo(0, 0)");

        // 9) Engranaje
        var gear = page.Locator("span.glyphicon.glyphicon-cog").First;
        await gear.WaitForAsync(new() { Timeout = 60000 });
        await gear.ClickAsync(new() { Force = true });

        // 10) Cerrar sesión
        await page.GetByRole(AriaRole.Link, new() { Name = "Cerrar Sesión" })
                  .ClickAsync(new() { Force = true });

        await page.WaitForSelectorAsync("#login", new() { Timeout = 120000 });
        Console.WriteLine("Sesión cerrada correctamente.");

        // Leer Excel con tu clase (Services/VeritradeExcelReader.cs)
        if (!string.IsNullOrWhiteSpace(excelPath) && File.Exists(excelPath))
        {
            var reader = new VeritradeExcelReader();
            var rows = reader.Read(excelPath);

            var repo = new VeritradeSqlRepository(cs);
            var cargaId = Guid.NewGuid().ToString("N");           // único por excel
            var cargaFecha = DateTime.UtcNow;       // fecha/hora de carga (recomendado UTC)
            await repo.InsertAsync(rows, cargaId, cargaFecha);

            Console.WriteLine("Insertado en dbo.ImportacionesAduanas (Server25).");


            Console.WriteLine($"Filas leídas: {rows.Count}");

            foreach (var r in rows)
            {
                Console.WriteLine($"Partida: {r.PartidaAduanera} | Desc: {r.DescripcionPartidaAduanera} | Importador: {r.Importador}");
            }
        }
        else
        {
            Console.WriteLine("No se encontró el archivo Excel para leer.");
        }

        // (Opcional) cerrar navegador
        await browser.CloseAsync();
    }
}


/*using Microsoft.Playwright;
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
            "Descargas"
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
}*/


/*
using Microsoft.Playwright;
using System;
using System.IO;
using System.Threading.Tasks;
using ExtraeData.Services;
using ExtraeData.Data;
using Microsoft.Extensions.Configuration;
using System.Globalization;

class Program
{
    static async Task Main()
    {
        // ===== SQL =====
        var config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)    
            .AddJsonFile("appsettings.json", optional: false)
            .Build();

        var cs = config.GetConnectionString("Server25")
                 ?? throw new Exception("Falta ConnectionStrings:Server25 en appsettings.json");

        // ===== Credenciales =====
        var user = "rossana.iglesias@nogasa.com.pe";
        var pass = "Vistony2020";

        // ===== Destino de descarga =====
        var downloadsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Descargas"
        );
        Directory.CreateDirectory(downloadsDir);

        //LISTA DE PARTIDAS PERU-IMPOTACION
        var partidas = new[] {

            "3811211000",
            "2710193800",
            "3403990000",
            "2710193400",
            "2710193600",
            "3403190000",
            "2710193500",
            "3819000000",
            "3811219000",
            "3811211000-",
            "3823190000-",
            "3811290000",
            "2710193900",
            "3902900000",
            "2830909000",
            "3811900000",
            "3820000000",
            "2905310000",
            "3102101000",
            "3901901000",
            "3811212000",
            "2905310000",
            "2825200000",
            "3910001000",
            "3901200000"
            
        };

        // CÓDIGOS PUNTUALES POR PAÍS (rellena con los que necesitas)
        var partidasChile = new[] {

            "2710193500",
            "29053100",
            "38231100",
            "2825200000",
            "39100090",
            "310210",
            "3901200000",
            "39100090"
        };
        var partidasColombia = new[] {    

            "2710190000",
            "2905310000",
            "382311",
            "2825200000",
            "3910001000",
            "3102101000",
            "3901200000"
        };
        var partidasEcuador = new[] {
        
            "2905310000",
            "3823110000",
            "2825200000",
            "3910001000",
            "3102101000",
            "3901200000"
        };

        var partidasExportPeru = new[] {
            "2710193800" };


        //time de proceso

        using var playwright = await Playwright.CreateAsync();

        var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
        {
            Headless = false,
            SlowMo = 90,
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
            ViewportSize = null, //para que no se vea “pequeño”
            AcceptDownloads = true,
            UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        });

        //Quitar webdriver (igual que tu versión que funcionaba)
        await context.AddInitScriptAsync("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});");

        var page = await context.NewPageAsync();

        await page.GotoAsync("https://www.veritradecorp.com/",
            new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded, Timeout = 120000 });

        // =========================
        // LOGIN
        // =========================
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

        // Espera robusta a que estés dentro (panel)  
        await page.WaitForURLAsync("**es/mis-busquedas**", new() { Timeout = 120000 }); //quitamis un "/ "**es/
        await page.WaitForSelectorAsync("#btnBuscar", new() { Timeout = 120000 });

        // ===== Reader/Repo =====
        var reader = new VeritradeExcelReader();
        var repo = new VeritradeSqlRepository(cs);

        // =========================
        // 2) PROCESO (LOOP) PPERU
        // =========================
        for (int i = 0; i < partidas.Length; i++)
        {
            var partida = partidas[i];
            Console.WriteLine($"\n=== Partida {partida} ({i + 1}/{partidas.Length}) ===");

            // 2.1 Asegurar “Importaciones”
            var rdbImp = page.Locator("#rdbImp");
            if (await rdbImp.CountAsync() > 0 && !await rdbImp.IsCheckedAsync())
                await rdbImp.CheckAsync(new() { Force = true });

            // 2.2 Asegurar “Perú”
            var cboPais = page.Locator("#cboPais");
            if (await cboPais.CountAsync() > 0)
            {
                try { await cboPais.SelectOptionAsync(new SelectOptionValue { Label = "Perú" }); }
                catch { await cboPais.SelectOptionAsync(new SelectOptionValue { Label = "Peru" }); }
            }

            // ===== PARTIDA ADUANERA (forzar tipeo real) ====
            // 2.3 Partida Aduanera: escribir + seleccionar sugerencia y confirmar que quedó en "Filtros"
            // ===== PARTIDA ADUANERA: escribir → esperar lista → DOBLE CLICK → validar en #lstFiltros =====
            var partidaInput = page.Locator("#txtNandinaB");
            await partidaInput.WaitForAsync(new() { Timeout = 60000 });

            async Task TypePartidaAsync(string partida)
            {
                await partidaInput.ClickAsync(new() { Force = true });
                await partidaInput.PressAsync("Control+A");
                await partidaInput.PressAsync("Backspace");

                // tipeo humano + eventos reales
                await partidaInput.TypeAsync(partida, new() { Delay = 80 });

                // dispara keyup extra (a veces jQuery UI lo necesita)
                await partidaInput.PressAsync("ArrowRight");
                await page.WaitForTimeoutAsync(200);
            }

            // reintenta si el input se “limpia”
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                await TypePartidaAsync(partida);

                // espera breve a que JS procese y verifica que el valor quedó
                await page.WaitForTimeoutAsync(400);
                var current = await partidaInput.InputValueAsync();

                if (!string.IsNullOrWhiteSpace(current) && current.Contains(partida))
                    break;

                if (attempt == 3)
                    throw new Exception($"No se pudo mantener el valor '{partida}' en #txtNandinaB (se borra).");
            }

            // ===== Esperar autocomplete =====
            // ===== Seleccionar sugerencia del autocomplete =====
            // Espera el item visible que contenga la partida (o el primero visible)
            var itemExacto = page.Locator($"li.ui-menu-item:visible >> text={partida}").First;

            // a veces el texto está dentro de <div> o <a>, este selector igual lo encuentra
            if (await itemExacto.CountAsync() == 0)
                itemExacto = page.Locator("li.ui-menu-item:visible").First;

            await itemExacto.WaitForAsync(new() { Timeout = 60000 });

            // ✅ Doble click como manual
            await itemExacto.DblClickAsync(new() { Force = true });

            // Validar que se agregó al select de filtros (#lstFiltros)
            await page.WaitForFunctionAsync(
                @"(p) => {
        const sel = document.querySelector('#lstFiltros');
        if (!sel) return false;
        return Array.from(sel.options || []).some(o => (o.textContent||'').includes(p));
    }",
                partida,
                new() { Timeout = 60000 }
            );

            // ===== FECHAS: poner MES ANTERIOR (ej: febrero si estamos en marzo) =====
            var prev = DateTime.Today.AddMonths(-1);

            // UI usa: Ene Feb Mar Abr May Jun Jul Ago Sep Oct Nov Dic
            string[] meses = { "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic" };
            var mesUi = meses[prev.Month - 1];
            var esperado = $"{mesUi} {prev:yyyy}";

            async Task PickPrevMonthAsync(ILocator input)
            {
                await input.WaitForAsync(new() { Timeout = 60000 });

                // abrir datepicker: primero intenta click en input, si no abre, click en el icono/calendario del mismo control
                await input.ClickAsync(new() { Force = true });

                // espera corta a ver si aparece el datepicker
                var picker = page.Locator(".datepicker:visible, .datepicker-dropdown:visible").First;
                try
                {
                    await picker.WaitForAsync(new() { Timeout = 1500 });
                }
                catch
                {
                    // no abrió con click al input -> click al botón/icono al costado (input-group)
                    var iconBtn = input.Locator("xpath=ancestor::*[contains(@class,'input-group')][1]//button").First;
                    if (await iconBtn.CountAsync() == 0)
                        iconBtn = input.Locator("xpath=ancestor::*[contains(@class,'input-group')][1]//span[contains(@class,'input-group-addon') or contains(@class,'add-on')]").First;

                    if (await iconBtn.CountAsync() > 0)
                        await iconBtn.ClickAsync(new() { Force = true });
                    else
                        // último recurso: el hermano siguiente (muchas UIs ponen el botón a la derecha)
                        await input.Locator("xpath=following-sibling::*[1]").ClickAsync(new() { Force = true });

                    // ahora sí espera el datepicker visible
                    picker = page.Locator(".datepicker:visible, .datepicker-dropdown:visible").First;
                    await picker.WaitForAsync(new() { Timeout = 60000 });
                }

                // pasar a vista "meses" si está en días
                var monthsView = picker.Locator(".datepicker-months");
                if (await monthsView.CountAsync() > 0 && !await monthsView.IsVisibleAsync())
                {
                    var sw = picker.Locator("th.datepicker-switch").First;
                    // 1 click: días -> meses (en la mayoría de configs)
                    await sw.ClickAsync(new() { Force = true });

                    // si aún no aparece, intenta una vez más (por si entra a años primero)
                    if (!await monthsView.IsVisibleAsync())
                        await sw.ClickAsync(new() { Force = true });
                }

                await monthsView.WaitForAsync(new() { Timeout = 60000 });

                // click al mes anterior (no disabled)
                var mesBtn = picker.Locator($".datepicker-months span.month:not(.disabled)", new() { HasTextString = mesUi }).First;
                await mesBtn.WaitForAsync(new() { Timeout = 60000 });
                await mesBtn.ClickAsync(new() { Force = true });

                // esperar a que el input refleje el cambio (que no se quede en Mar 2026)
                var handle = await input.ElementHandleAsync();
                await page.WaitForFunctionAsync(
                    @"(p) => {
            const el = p.el;
            const expected = p.expected;
            return el && (el.value || '').includes(expected);
        }",
                    new { el = handle, expected = esperado },
                    new() { Timeout = 60000 }
                );
            }

            // Desde
            var fromInput = page.Locator("#cboDesde input").First;
            await PickPrevMonthAsync(fromInput);

            // Hasta
            var toInput = page.Locator("#cboHasta input").First;
            await PickPrevMonthAsync(toInput);

            // log de confirmación
            Console.WriteLine($"Desde={(await fromInput.InputValueAsync()).Trim()} | Hasta={(await toInput.InputValueAsync()).Trim()} | Esperado={esperado}");
         
            /////////////buscar///////////////
            // 2.5 Buscar
            await page.Locator("#btnBuscar").ClickAsync(new() { Force = true });

            // 2.5.1 Esperar un poquito para que aparezca modal o resultados
            await page.WaitForTimeoutAsync(800);

            // Si sale modal de "no encontró resultados"
            var noResultsOk = page.Locator("#btnOKModalVentanaMensaje");
            if (await noResultsOk.CountAsync() > 0 && await noResultsOk.IsVisibleAsync())
            {
                Console.WriteLine($"[{partida}] No se encontraron registros. Aceptar + Restablecer + siguiente código.");

                // 1) ACEPTAR
                await noResultsOk.ClickAsync(new() { Force = true });

                // 2) esperar cierre modal
                await page.WaitForTimeoutAsync(800);

                // 3) RESTABLECER
                await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });

                // 4) esperar limpieza
                await page.WaitForTimeoutAsync(1200);

                //siguiente código (solo si estás dentro del foreach)
                continue;
            }

            // 2.5.2 Si NO hubo modal: esperar resultados (sin Options Timeout)
            var start = DateTime.UtcNow;
            while (true)
            {
                // si en algún momento aparece el modal, lo manejamos también
                if (await noResultsOk.CountAsync() > 0 && await noResultsOk.IsVisibleAsync())
                {
                    Console.WriteLine($"[{partida}] No se encontraron registros (tardío). Aceptar + Restablecer + siguiente código.");

                    await noResultsOk.ClickAsync(new() { Force = true });
                    await page.WaitForTimeoutAsync(800);

                    await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });
                    await page.WaitForTimeoutAsync(1200);

                    continue; // siguiente partida
                }

                // label de resultados
                var total = page.Locator("#totalRecordsFound");
                if (await total.CountAsync() > 0)
                {
                    var t = (await total.InnerTextAsync()).Trim();
                    if (t.IndexOf("Se encontraron", StringComparison.OrdinalIgnoreCase) >= 0 &&
                        t.IndexOf("registros", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine($"[{partida}] OK => {t}");
                        break;
                    }
                }

                // timeout manual: 3 min
                if ((DateTime.UtcNow - start).TotalMilliseconds > 180000)
                    throw new TimeoutException($"[{partida}] Timeout esperando '#totalRecordsFound' con 'Se encontraron ... registros'.");

                await page.WaitForTimeoutAsync(500);
            }


            // 2.6 Ir a “Detalle - Excel”      
            await page.Locator("#tabDetalleExcel").ClickAsync(new() { Force = true });

            // espera que el botón exista (aunque esté oculto un rato)
            var excelBtn = page.Locator("#downloadFileVerRegistro2");
            await excelBtn.WaitForAsync(new() { Timeout = 120000, State = WaitForSelectorState.Attached });

            // a veces queda oculto/disabled mientras carga la tabla, espera a que pueda clickease
            await page.WaitForFunctionAsync(
                        @"() => {
                const b = document.querySelector('#downloadFileVerRegistro2');
                if (!b) return false;
                const style = window.getComputedStyle(b);
                const visible = style && style.visibility !== 'hidden' && style.display !== 'none';
                const enabled = !b.hasAttribute('disabled');
                return visible && enabled;
            }",
                        null,
                new() { Timeout = 120000 }
            );

            // 2.7 Descargar Excel
            var downloadTask = page.WaitForDownloadAsync(new() { Timeout = 600000 }); // 10 min
            await excelBtn.ClickAsync(new() { Force = true });
            var download = await downloadTask;

            var filename = download.SuggestedFilename;
            if (string.IsNullOrWhiteSpace(filename))
                filename = "veritrade_detalle.xlsx";

            var fullPath = Path.Combine(downloadsDir, $"{DateTime.Now:yyyyMMdd_HHmmss}_{partida}_{filename}");
            await download.SaveAsAsync(fullPath);

            Console.WriteLine($"Excel descargado: {fullPath}");

            ////////////////////////////////////////////////
            ///////leer y cargar datos excel/////////////////
            ////////////////////////////////////////////////
            // 2.8 Leer Excel + Insertar SQL (cargaId único por Excel)        
            var rows = reader.Read(fullPath, "Perú");
            //log 1
            Console.WriteLine($"Excel leído: {Path.GetFileName(fullPath)} | rows.Count={rows.Count}");
            //log 2
            foreach (var r in rows.Take(3))
                Console.WriteLine($"[HEAD] {r.PartidaAduanera} | {r.DescripcionPartidaAduanera} | {r.Importador}");

            foreach (var r in rows.TakeLast(3))
                Console.WriteLine($"[TAIL] {r.PartidaAduanera} | {r.DescripcionPartidaAduanera} | {r.Importador}");
            //log 3
            var vacias = rows.Count(r =>
                string.IsNullOrWhiteSpace(r.PartidaAduanera) &&
                string.IsNullOrWhiteSpace(r.DescripcionPartidaAduanera) &&
                string.IsNullOrWhiteSpace(r.Importador));

            Console.WriteLine($"Filas totalmente vacías (3 campos): {vacias}");
            //log 3

            var cargaId = Guid.NewGuid().ToString("N");
            var cargaFecha = DateTime.UtcNow;
            //log 4
            Console.WriteLine($"Insertando en SQL... cargaId={cargaId} cargaFecha={cargaFecha:O}");


            await repo.InsertAsync(rows, cargaId, cargaFecha);
            //log 5
            Console.WriteLine($"Insertado SQL: filas(leídas)={rows.Count}, cargaId={cargaId}");


            Console.WriteLine($"Insertado SQL: filas={rows.Count}, cargaId={cargaId}");

            // 2.9 Restablecer buscador y siguiente
            await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });
            await page.WaitForTimeoutAsync(1200);

        }
        //FINAL DE LOOOP//


        // ===============================================================================================================================================
        // 2.B) NUEVO: DESCARGAS POR OTROS PAÍSES (CL/CO/EC)
        // =====================================================

        // helper: seleccionar país (primero intenta #cboPais2, si no existe usa #cboPais)
        async Task SelectCountryAsync(string label)
        {
            // ✅ Este es el combo correcto donde aparece "Perú" y la lista de países
            var cbo = page.Locator("#cboPais");
            if (await cbo.CountAsync() == 0)
                cbo = page.Locator("#cboPais2"); // fallback por si en alguna vista cambia

            await cbo.WaitForAsync(new() { Timeout = 60000 });

            var handle = await cbo.ElementHandleAsync();

            // esperar opciones cargadas
            await page.WaitForFunctionAsync(
                @"(sel) => sel && sel.options && sel.options.length > 1",
                handle,
                new() { Timeout = 60000 }
            );

            // seleccionar por texto (contiene)
            var ok = await page.EvaluateAsync<bool>(
                @"(p) => {
            const sel = p.sel;
            const txt = p.txt;
            const norm = s => (s||'').trim().toLowerCase();
            const t = norm(txt);
            const opts = Array.from(sel.options || []);
            const hit = opts.find(o => norm(o.textContent).includes(t));
            if (!hit) return false;
            sel.value = hit.value;
            sel.dispatchEvent(new Event('input', { bubbles: true }));
            sel.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }",
                new { sel = handle, txt = label }
            );

            if (!ok)
                throw new Exception($"No encontré '{label}' dentro del selector de países (#cboPais).");

            await page.WaitForTimeoutAsync(800);
        }


        // helper: ejecutar proceso por país
        async Task RunForCountryAsync(string countryLabel, string[] countryPartidas)
        {
            if (countryPartidas == null || countryPartidas.Length == 0)
            {
                Console.WriteLine($"⚠️ {countryLabel}: no hay partidas definidas, se omite.");
                return;
            }

            Console.WriteLine($"\n===== INICIANDO PAÍS: {countryLabel} | partidas={countryPartidas.Length} =====");

            // Asegurar “Importaciones”
            var rdbImp = page.Locator("#rdbImp");
            if (await rdbImp.CountAsync() > 0 && !await rdbImp.IsCheckedAsync())
                await rdbImp.CheckAsync(new() { Force = true });

            // seleccionar el país (Chile/Colombia/Ecuador)
            await SelectCountryAsync(countryLabel);

            // recorrer SOLO las partidas de ese país
            for (int j = 0; j < countryPartidas.Length; j++)
            {
                var partida = countryPartidas[j];
                Console.WriteLine($"\n=== [{countryLabel}] Partida {partida} ({j + 1}/{countryPartidas.Length}) ===");

                var partidaInput = page.Locator("#txtNandinaB");
                await partidaInput.WaitForAsync(new() { Timeout = 60000 });

                async Task TypePartidaAsync(string partida)
                {
                    await partidaInput.ClickAsync(new() { Force = true });
                    await partidaInput.PressAsync("Control+A");
                    await partidaInput.PressAsync("Backspace");

                    // tipeo humano + eventos reales
                    await partidaInput.TypeAsync(partida, new() { Delay = 80 });

                    // dispara keyup extra (a veces jQuery UI lo necesita)
                    await partidaInput.PressAsync("ArrowRight");
                    await page.WaitForTimeoutAsync(200);
                }

                // reintenta si el input se “limpia”
                for (int attempt = 1; attempt <= 3; attempt++)
                {
                    await TypePartidaAsync(partida);

                    // espera breve a que JS procese y verifica que el valor quedó
                    await page.WaitForTimeoutAsync(400);
                    var current = await partidaInput.InputValueAsync();

                    if (!string.IsNullOrWhiteSpace(current) && current.Contains(partida))
                        break;

                    if (attempt == 3)
                        throw new Exception($"No se pudo mantener el valor '{partida}' en #txtNandinaB (se borra).");
                }

                // ===== Esperar autocomplete =====
                // ===== Seleccionar sugerencia del autocomplete =====
                // Espera el item visible que contenga la partida (o el primero visible)
                var itemExacto = page.Locator($"li.ui-menu-item:visible >> text={partida}").First;

                // a veces el texto está dentro de <div> o <a>, este selector igual lo encuentra
                if (await itemExacto.CountAsync() == 0)
                    itemExacto = page.Locator("li.ui-menu-item:visible").First;

                await itemExacto.WaitForAsync(new() { Timeout = 60000 });

                // Doble click como manual
                await itemExacto.DblClickAsync(new() { Force = true });

                // Validar que se agregó al select de filtros (#lstFiltros)
                await page.WaitForFunctionAsync(
                    @"(p) => {
        const sel = document.querySelector('#lstFiltros');
        if (!sel) return false;
        return Array.from(sel.options || []).some(o => (o.textContent||'').includes(p));
    }",
                    partida,
                    new() { Timeout = 60000 }
                );

                // ===== FECHAS: poner MES ANTERIOR (ej: febrero si estamos en marzo) =====
                var prev = DateTime.Today.AddMonths(-1);

                // UI usa: Ene Feb Mar Abr May Jun Jul Ago Sep Oct Nov Dic
                string[] meses = { "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic" };
                var mesUi = meses[prev.Month - 1];
                var esperado = $"{mesUi} {prev:yyyy}";

                async Task PickPrevMonthAsync(ILocator input)
                {
                    await input.WaitForAsync(new() { Timeout = 60000 });

                    // abrir datepicker: primero intenta click en input, si no abre, click en el icono/calendario del mismo control
                    await input.ClickAsync(new() { Force = true });

                    // espera corta a ver si aparece el datepicker
                    var picker = page.Locator(".datepicker:visible, .datepicker-dropdown:visible").First;
                    try
                    {
                        await picker.WaitForAsync(new() { Timeout = 1500 });
                    }
                    catch
                    {
                        // no abrió con click al input -> click al botón/icono al costado (input-group)
                        var iconBtn = input.Locator("xpath=ancestor::*[contains(@class,'input-group')][1]//button").First;
                        if (await iconBtn.CountAsync() == 0)
                            iconBtn = input.Locator("xpath=ancestor::*[contains(@class,'input-group')][1]//span[contains(@class,'input-group-addon') or contains(@class,'add-on')]").First;

                        if (await iconBtn.CountAsync() > 0)
                            await iconBtn.ClickAsync(new() { Force = true });
                        else
                            // último recurso: el hermano siguiente (muchas UIs ponen el botón a la derecha)
                            await input.Locator("xpath=following-sibling::*[1]").ClickAsync(new() { Force = true });

                        // ahora sí espera el datepicker visible
                        picker = page.Locator(".datepicker:visible, .datepicker-dropdown:visible").First;
                        await picker.WaitForAsync(new() { Timeout = 60000 });
                    }

                    // pasar a vista "meses" si está en días
                    var monthsView = picker.Locator(".datepicker-months");
                    if (await monthsView.CountAsync() > 0 && !await monthsView.IsVisibleAsync())
                    {
                        var sw = picker.Locator("th.datepicker-switch").First;
                        // 1 click: días -> meses (en la mayoría de configs)
                        await sw.ClickAsync(new() { Force = true });

                        // si aún no aparece, intenta una vez más (por si entra a años primero)
                        if (!await monthsView.IsVisibleAsync())
                            await sw.ClickAsync(new() { Force = true });
                    }

                    await monthsView.WaitForAsync(new() { Timeout = 60000 });

                    // click al mes anterior (no disabled)
                    var mesBtn = picker.Locator($".datepicker-months span.month:not(.disabled)", new() { HasTextString = mesUi }).First;
                    await mesBtn.WaitForAsync(new() { Timeout = 60000 });
                    await mesBtn.ClickAsync(new() { Force = true });

                    // esperar a que el input refleje el cambio (que no se quede en Mar 2026)
                    var handle = await input.ElementHandleAsync();
                    await page.WaitForFunctionAsync(
                                    @"(p) => {
                        const el = p.el;
                        const expected = p.expected;
                        return el && (el.value || '').includes(expected);
                    }",
                        new { el = handle, expected = esperado },
                        new() { Timeout = 60000 }
                    );
                }

            
                // ===== FECHAS: para CL/CO/EC NO tocar, usar las fechas por defecto =====
                var fromInput = page.Locator("#cboDesde input").First;
                var toInput = page.Locator("#cboHasta input").First;

                // solo log de lo que viene por defecto
                Console.WriteLine($"[FECHAS DEFAULT {countryLabel}] Desde={(await fromInput.InputValueAsync()).Trim()} | Hasta={(await toInput.InputValueAsync()).Trim()}");

                /////////////buscar///////////////
                // 2.5 Buscar
                await page.Locator("#btnBuscar").ClickAsync(new() { Force = true });

                // 2.5.1 Esperar un poquito para que aparezca modal o resultados
                await page.WaitForTimeoutAsync(800);

                // Si sale modal de "no encontró resultados"
                var noResultsOk = page.Locator("#btnOKModalVentanaMensaje");
                if (await noResultsOk.CountAsync() > 0 && await noResultsOk.IsVisibleAsync())
                {
                    Console.WriteLine($"[{partida}] No se encontraron registros. Aceptar + Restablecer + siguiente código.");

                    // 1) ACEPTAR
                    await noResultsOk.ClickAsync(new() { Force = true });

                    // 2) esperar cierre modal
                    await page.WaitForTimeoutAsync(800);

                    // 3) RESTABLECER
                    await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });

                    // 4) esperar limpieza
                    await page.WaitForTimeoutAsync(1200);

                    //siguiente código (solo si estás dentro del foreach)
                    continue;
                }

                // 2.5.2 Si NO hubo modal: esperar resultados (sin Options Timeout)
                var start = DateTime.UtcNow;
                while (true)
                {
                    // si en algún momento aparece el modal, lo manejamos también
                    if (await noResultsOk.CountAsync() > 0 && await noResultsOk.IsVisibleAsync())
                    {
                        Console.WriteLine($"[{partida}] No se encontraron registros (tardío). Aceptar + Restablecer + siguiente código.");

                        await noResultsOk.ClickAsync(new() { Force = true });
                        await page.WaitForTimeoutAsync(800);

                        await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });
                        await page.WaitForTimeoutAsync(1200);

                        continue; // siguiente partida
                    }

                    // label de resultados
                    var total = page.Locator("#totalRecordsFound");
                    if (await total.CountAsync() > 0)
                    {
                        var t = (await total.InnerTextAsync()).Trim();
                        if (t.IndexOf("Se encontraron", StringComparison.OrdinalIgnoreCase) >= 0 &&
                            t.IndexOf("registros", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine($"[{partida}] OK => {t}");
                            break;
                        }
                    }

                    // timeout manual: 3 min
                    if ((DateTime.UtcNow - start).TotalMilliseconds > 180000)
                        throw new TimeoutException($"[{partida}] Timeout esperando '#totalRecordsFound' con 'Se encontraron ... registros'.");

                    await page.WaitForTimeoutAsync(800);
                }


                // 2.6 Ir a “Detalle - Excel”
                await page.Locator("#tabDetalleExcel").ClickAsync(new() { Force = true });

                // espera que el botón exista (aunque esté oculto un rato)
                var excelBtn = page.Locator("#downloadFileVerRegistro2");
                await excelBtn.WaitForAsync(new() { Timeout = 120000, State = WaitForSelectorState.Attached });

                // a veces queda oculto/disabled mientras carga la tabla, espera a que pueda clickease
                await page.WaitForFunctionAsync(
                                @"() => {
                    const b = document.querySelector('#downloadFileVerRegistro2');
                    if (!b) return false;
                    const style = window.getComputedStyle(b);
                    const visible = style && style.visibility !== 'hidden' && style.display !== 'none';
                    const enabled = !b.hasAttribute('disabled');
                    return visible && enabled;
                }",
                    null,
                    new() { Timeout = 120000 }
                );

                // 2.7 Descargar Excel
                var downloadTask = page.WaitForDownloadAsync(new() { Timeout = 600000 }); // 10 min
                await excelBtn.ClickAsync(new() { Force = true });
                var download = await downloadTask;

                var filename = download.SuggestedFilename;
                if (string.IsNullOrWhiteSpace(filename))
                    filename = "veritrade_detalle.xlsx";

                var fullPath = Path.Combine(downloadsDir, $"{DateTime.Now:yyyyMMdd_HHmmss}_{partida}_{filename}");
                await download.SaveAsAsync(fullPath);

                Console.WriteLine($"Excel descargado: {fullPath}");

                ////////////////////////////////////////////////
                ///////leer y cargar datos excel/////////////////
                ////////////////////////////////////////////////
                // 2.8 Leer Excel + Insertar SQL (cargaId único por Excel)

                var rows = reader.Read(fullPath, countryLabel);
                //log 1
                Console.WriteLine($"Excel leído: {Path.GetFileName(fullPath)} | rows.Count={rows.Count}");
                //log 2
                foreach (var r in rows.Take(3))
                    Console.WriteLine($"[HEAD] {r.PartidaAduanera} | {r.DescripcionPartidaAduanera} | {r.Importador}");

                foreach (var r in rows.TakeLast(3))
                    Console.WriteLine($"[TAIL] {r.PartidaAduanera} | {r.DescripcionPartidaAduanera} | {r.Importador}");
                //log 3
                var vacias = rows.Count(r =>
                    string.IsNullOrWhiteSpace(r.PartidaAduanera) &&
                    string.IsNullOrWhiteSpace(r.DescripcionPartidaAduanera) &&
                    string.IsNullOrWhiteSpace(r.Importador));

                Console.WriteLine($"Filas totalmente vacías (3 campos): {vacias}");
                //log 3

                var cargaId = Guid.NewGuid().ToString("N");
                var cargaFecha = DateTime.UtcNow;
                //log 4
                Console.WriteLine($"Insertando en SQL... cargaId={cargaId} cargaFecha={cargaFecha:O}");


                await repo.InsertAsync(rows, cargaId, cargaFecha);
                //log 5
                Console.WriteLine($"Insertado SQL: filas(leídas)={rows.Count}, cargaId={cargaId}");


                Console.WriteLine($"Insertado SQL: filas={rows.Count}, cargaId={cargaId}");

                // 2.9 Restablecer buscador y siguiente
                await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });
                await page.WaitForTimeoutAsync(1200);

            }
        
        }

        // Ejecutar en orden (DESPUÉS de definir las funciones)
        await RunForCountryAsync("Chile", partidasChile);
        await RunForCountryAsync("Colombia", partidasColombia);
        await RunForCountryAsync("Ecuador", partidasEcuador);



        // =========================
        // 2.C) EXPORTACIONES - PERÚ (solo 2710193800)
        // =========================
        for (int i = 0; i < partidasExportPeru.Length; i++)
        {
            var partida = partidasExportPeru[i];
            Console.WriteLine($"\n=== [EXPORT PERÚ] Partida {partida} ({i + 1}/{partidasExportPeru.Length}) ===");

            // ✅ 1) Asegurar “Exportaciones”
            var rdbExp = page.Locator("#rdbExp");
            if (await rdbExp.CountAsync() > 0 && !await rdbExp.IsCheckedAsync())
                await rdbExp.CheckAsync(new() { Force = true });

            // ✅ 2) Asegurar “Perú” (combo de país)
            var cboPais = page.Locator("#cboPais");
            if (await cboPais.CountAsync() > 0)
            {
                try { await cboPais.SelectOptionAsync(new SelectOptionValue { Label = "Perú" }); }
                catch { await cboPais.SelectOptionAsync(new SelectOptionValue { Label = "Peru" }); }
            }

            var partidaInput = page.Locator("#txtNandinaB");
            await partidaInput.WaitForAsync(new() { Timeout = 60000 });

            async Task TypePartidaAsync(string partida)
            {
                await partidaInput.ClickAsync(new() { Force = true });
                await partidaInput.PressAsync("Control+A");
                await partidaInput.PressAsync("Backspace");

                // tipeo humano + eventos reales
                await partidaInput.TypeAsync(partida, new() { Delay = 80 });

                // dispara keyup extra (a veces jQuery UI lo necesita)
                await partidaInput.PressAsync("ArrowRight");
                await page.WaitForTimeoutAsync(200);
            }

            // reintenta si el input se “limpia”
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                await TypePartidaAsync(partida);

                // espera breve a que JS procese y verifica que el valor quedó
                await page.WaitForTimeoutAsync(400);
                var current = await partidaInput.InputValueAsync();

                if (!string.IsNullOrWhiteSpace(current) && current.Contains(partida))
                    break;

                if (attempt == 3)
                    throw new Exception($"No se pudo mantener el valor '{partida}' en #txtNandinaB (se borra).");
            }

            // ===== Esperar autocomplete =====
            // ===== Seleccionar sugerencia del autocomplete =====
            // Espera el item visible que contenga la partida (o el primero visible)
            var itemExacto = page.Locator($"li.ui-menu-item:visible >> text={partida}").First;

            // a veces el texto está dentro de <div> o <a>, este selector igual lo encuentra
            if (await itemExacto.CountAsync() == 0)
                itemExacto = page.Locator("li.ui-menu-item:visible").First;

            await itemExacto.WaitForAsync(new() { Timeout = 60000 });

            // ✅ Doble click como manual
            await itemExacto.DblClickAsync(new() { Force = true });

            // Validar que se agregó al select de filtros (#lstFiltros)
            await page.WaitForFunctionAsync(
                @"(p) => {
        const sel = document.querySelector('#lstFiltros');
        if (!sel) return false;
        return Array.from(sel.options || []).some(o => (o.textContent||'').includes(p));
    }",
                partida,
                new() { Timeout = 60000 }
            );

            // ===== FECHAS: poner MES ANTERIOR (ej: febrero si estamos en marzo) =====
            var prev = DateTime.Today.AddMonths(-1);

            // UI usa: Ene Feb Mar Abr May Jun Jul Ago Sep Oct Nov Dic
            string[] meses = { "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic" };
            var mesUi = meses[prev.Month - 1];
            var esperado = $"{mesUi} {prev:yyyy}";

            async Task PickPrevMonthAsync(ILocator input)
            {
                await input.WaitForAsync(new() { Timeout = 60000 });

                // abrir datepicker: primero intenta click en input, si no abre, click en el icono/calendario del mismo control
                await input.ClickAsync(new() { Force = true });

                // espera corta a ver si aparece el datepicker
                var picker = page.Locator(".datepicker:visible, .datepicker-dropdown:visible").First;
                try
                {
                    await picker.WaitForAsync(new() { Timeout = 1500 });
                }
                catch
                {
                    // no abrió con click al input -> click al botón/icono al costado (input-group)
                    var iconBtn = input.Locator("xpath=ancestor::*[contains(@class,'input-group')][1]//button").First;
                    if (await iconBtn.CountAsync() == 0)
                        iconBtn = input.Locator("xpath=ancestor::*[contains(@class,'input-group')][1]//span[contains(@class,'input-group-addon') or contains(@class,'add-on')]").First;

                    if (await iconBtn.CountAsync() > 0)
                        await iconBtn.ClickAsync(new() { Force = true });
                    else
                        // último recurso: el hermano siguiente (muchas UIs ponen el botón a la derecha)
                        await input.Locator("xpath=following-sibling::*[1]").ClickAsync(new() { Force = true });

                    // ahora sí espera el datepicker visible
                    picker = page.Locator(".datepicker:visible, .datepicker-dropdown:visible").First;
                    await picker.WaitForAsync(new() { Timeout = 60000 });
                }

                // pasar a vista "meses" si está en días
                var monthsView = picker.Locator(".datepicker-months");
                if (await monthsView.CountAsync() > 0 && !await monthsView.IsVisibleAsync())
                {
                    var sw = picker.Locator("th.datepicker-switch").First;
                    // 1 click: días -> meses (en la mayoría de configs)
                    await sw.ClickAsync(new() { Force = true });

                    // si aún no aparece, intenta una vez más (por si entra a años primero)
                    if (!await monthsView.IsVisibleAsync())
                        await sw.ClickAsync(new() { Force = true });
                }

                await monthsView.WaitForAsync(new() { Timeout = 60000 });

                // click al mes anterior (no disabled)
                var mesBtn = picker.Locator($".datepicker-months span.month:not(.disabled)", new() { HasTextString = mesUi }).First;
                await mesBtn.WaitForAsync(new() { Timeout = 60000 });
                await mesBtn.ClickAsync(new() { Force = true });

                // esperar a que el input refleje el cambio (que no se quede en Mar 2026)
                var handle = await input.ElementHandleAsync();
                await page.WaitForFunctionAsync(
                    @"(p) => {
            const el = p.el;
            const expected = p.expected;
            return el && (el.value || '').includes(expected);
        }",
                    new { el = handle, expected = esperado },
                    new() { Timeout = 60000 }
                );
            }

            // Desde
            var fromInput = page.Locator("#cboDesde input").First;
            await PickPrevMonthAsync(fromInput);

            // Hasta
            var toInput = page.Locator("#cboHasta input").First;
            await PickPrevMonthAsync(toInput);

            // log de confirmación
            Console.WriteLine($"Desde={(await fromInput.InputValueAsync()).Trim()} | Hasta={(await toInput.InputValueAsync()).Trim()} | Esperado={esperado}");

            /////////////buscar///////////////
            // 2.5 Buscar
            await page.Locator("#btnBuscar").ClickAsync(new() { Force = true });

            // 2.5.1 Esperar un poquito para que aparezca modal o resultados
            await page.WaitForTimeoutAsync(800);

            // Si sale modal de "no encontró resultados"
            var noResultsOk = page.Locator("#btnOKModalVentanaMensaje");
            if (await noResultsOk.CountAsync() > 0 && await noResultsOk.IsVisibleAsync())
            {
                Console.WriteLine($"[{partida}] No se encontraron registros. Aceptar + Restablecer + siguiente código.");

                // 1) ACEPTAR
                await noResultsOk.ClickAsync(new() { Force = true });

                // 2) esperar cierre modal
                await page.WaitForTimeoutAsync(800);

                // 3) RESTABLECER
                await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });

                // 4) esperar limpieza
                await page.WaitForTimeoutAsync(1200);

                //siguiente código (solo si estás dentro del foreach)
                continue;
            }

            // 2.5.2 Si NO hubo modal: esperar resultados (sin Options Timeout)
            var start = DateTime.UtcNow;
            while (true)
            {
                // si en algún momento aparece el modal, lo manejamos también
                if (await noResultsOk.CountAsync() > 0 && await noResultsOk.IsVisibleAsync())
                {
                    Console.WriteLine($"[{partida}] No se encontraron registros (tardío). Aceptar + Restablecer + siguiente código.");

                    await noResultsOk.ClickAsync(new() { Force = true });
                    await page.WaitForTimeoutAsync(800);

                    await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });
                    await page.WaitForTimeoutAsync(1200);

                    continue; // siguiente partida
                }

                // label de resultados
                var total = page.Locator("#totalRecordsFound");
                if (await total.CountAsync() > 0)
                {
                    var t = (await total.InnerTextAsync()).Trim();
                    if (t.IndexOf("Se encontraron", StringComparison.OrdinalIgnoreCase) >= 0 &&
                        t.IndexOf("registros", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine($"[{partida}] OK => {t}");
                        break;
                    }
                }

                // timeout manual: 3 min
                if ((DateTime.UtcNow - start).TotalMilliseconds > 180000)
                    throw new TimeoutException($"[{partida}] Timeout esperando '#totalRecordsFound' con 'Se encontraron ... registros'.");

                await page.WaitForTimeoutAsync(500);
            }


            // 2.6 Ir a “Detalle - Excel”
            await page.Locator("#tabDetalleExcel").ClickAsync(new() { Force = true });

            // espera que el botón exista (aunque esté oculto un rato)
            var excelBtn = page.Locator("#downloadFileVerRegistro2");
            await excelBtn.WaitForAsync(new() { Timeout = 120000, State = WaitForSelectorState.Attached });

            // a veces queda oculto/disabled mientras carga la tabla, espera a que pueda clickease
            await page.WaitForFunctionAsync(
                        @"() => {
                const b = document.querySelector('#downloadFileVerRegistro2');
                if (!b) return false;
                const style = window.getComputedStyle(b);
                const visible = style && style.visibility !== 'hidden' && style.display !== 'none';
                const enabled = !b.hasAttribute('disabled');
                return visible && enabled;
            }",
                        null,
                new() { Timeout = 120000 }
            );

            // 2.7 Descargar Excel
            var downloadTask = page.WaitForDownloadAsync(new() { Timeout = 600000 }); // 10 min
            await excelBtn.ClickAsync(new() { Force = true });
            var download = await downloadTask;

            var filename = download.SuggestedFilename;
            if (string.IsNullOrWhiteSpace(filename))
                filename = "veritrade_detalle.xlsx";

            var fullPath = Path.Combine(downloadsDir, $"{DateTime.Now:yyyyMMdd_HHmmss}_{partida}_{filename}");
            await download.SaveAsAsync(fullPath);

            Console.WriteLine($"Excel descargado: {fullPath}");

            ////////////////////////////////////////////////
            ///////leer y cargar datos excel/////////////////
            ////////////////////////////////////////////////
            // 2.8 Leer Excel + Insertar SQL (cargaId único por Excel)
            var rows = reader.Read(fullPath, "Perú");
            //log 1
            Console.WriteLine($"Excel leído: {Path.GetFileName(fullPath)} | rows.Count={rows.Count}");
            //log 2
            foreach (var r in rows.Take(3))
                Console.WriteLine($"[HEAD] {r.PartidaAduanera} | {r.DescripcionPartidaAduanera} | {r.Importador}");

            foreach (var r in rows.TakeLast(3))
                Console.WriteLine($"[TAIL] {r.PartidaAduanera} | {r.DescripcionPartidaAduanera} | {r.Importador}");
            //log 3
            var vacias = rows.Count(r =>
                string.IsNullOrWhiteSpace(r.PartidaAduanera) &&
                string.IsNullOrWhiteSpace(r.DescripcionPartidaAduanera) &&
                string.IsNullOrWhiteSpace(r.Importador));

            Console.WriteLine($"Filas totalmente vacías (3 campos): {vacias}");
            //log 3

            var cargaId = Guid.NewGuid().ToString("N");
            var cargaFecha = DateTime.UtcNow;
            //log 4
            Console.WriteLine($"Insertando en SQL... cargaId={cargaId} cargaFecha={cargaFecha:O}");


            await repo.InsertAsync(rows, cargaId, cargaFecha);
            //log 5
            Console.WriteLine($"Insertado SQL: filas(leídas)={rows.Count}, cargaId={cargaId}");


            Console.WriteLine($"Insertado SQL: filas={rows.Count}, cargaId={cargaId}");

            // 2.9 Restablecer buscador y siguiente
            await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });
            await page.WaitForTimeoutAsync(1200);

    }



        // =========================
        // 3) Cerrar sesión
        // =========================
        await page.EvaluateAsync("window.scrollTo(0, 0)");

        var gear = page.Locator("span.glyphicon.glyphicon-cog").First;
        await gear.WaitForAsync(new() { Timeout = 60000 });
        await gear.ClickAsync(new() { Force = true });

        await page.GetByRole(AriaRole.Link, new() { Name = "Cerrar Sesión" })
                  .ClickAsync(new() { Force = true });

        await page.WaitForSelectorAsync("#login", new() { Timeout = 120000 });
        Console.WriteLine("Sesión cerrada correctamente.");

        await browser.CloseAsync();
    }
}

*/

using System.Threading.Tasks;
using ExtraeData.Rpa;

class Program
{
    static async Task Main()
    {
        await VeritradeRunner.RunAsync();
    }
}
using ExtraeData.Constants;
using ExtraeData.Data;
using ExtraeData.Models;
using ExtraeData.Services;
using Microsoft.Playwright;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExtraeData.Workflows
{
    public static class ExportacionesPeruWorkflow
    {
        public static async Task RunAsync(IPage page, string downloadsDir, VeritradeExcelReader reader, VeritradeSqlRepository repo)
        {
            var partidasExportPeru = Partidas.PeruExportaciones;

            for (int i = 0; i < partidasExportPeru.Length; i++)
            {
                var partida = partidasExportPeru[i];
                Console.WriteLine($"\n=== [EXPORT PERÚ] Partida {partida} ({i + 1}/{partidasExportPeru.Length}) ===");

                var rdbExp = page.Locator("#rdbExp");
                if (await rdbExp.CountAsync() > 0 && !await rdbExp.IsCheckedAsync())
                    await rdbExp.CheckAsync(new() { Force = true });

                var cboPais = page.Locator("#cboPais");
                if (await cboPais.CountAsync() > 0)
                {
                    try { await cboPais.SelectOptionAsync(new SelectOptionValue { Label = "Perú" }); }
                    catch { await cboPais.SelectOptionAsync(new SelectOptionValue { Label = "Peru" }); }
                }

                var partidaInput = page.Locator("#txtNandinaB");
                await partidaInput.WaitForAsync(new() { Timeout = 60000 });

                await page.WaitForFunctionAsync(
                                @"() => {
                    const el = document.querySelector('#txtNandinaB');
                    return el && !el.disabled;
                }",
                    null,
                    new() { Timeout = 30000 }
                );

                async Task TypePartidaAsync(string p)
                {
                    for (int attempt = 1; attempt <= 3; attempt++)
                    {
                        await page.WaitForFunctionAsync(
                            @"() => {
                const el = document.querySelector('#txtNandinaB');
                return el && !el.disabled;
                }",
                            null,
                            new() { Timeout = 30000 }
                        );

                        await partidaInput.ClickAsync();
                        await partidaInput.FillAsync("");
                        await page.WaitForTimeoutAsync(200);

                        await partidaInput.TypeAsync(p, new() { Delay = 80 });
                        await page.WaitForTimeoutAsync(800);

                        var current = await partidaInput.InputValueAsync();

                        if (!string.IsNullOrWhiteSpace(current) && current.Contains(p))
                        {
                            var suggestions = page.Locator("li.ui-menu-item:visible");

                            try
                            {
                                await suggestions.First.WaitForAsync(new() { Timeout = 2000 });
                            }
                            catch
                            {
                                // seguir reintentando
                            }

                            if (await suggestions.CountAsync() > 0)
                                return;
                        }

                        if (attempt < 3)
                        {
                            await page.WaitForTimeoutAsync(1000);
                        }
                        else
                        {
                            throw new Exception($"No se pudo escribir/mostrar autocomplete para '{p}' en #txtNandinaB.");
                        }
                    }
                }

                        await TypePartidaAsync(partida);



                var itemExacto = page.Locator($"li.ui-menu-item:visible >> text={partida}").First;
                if (await itemExacto.CountAsync() == 0)
                    itemExacto = page.Locator("li.ui-menu-item:visible").First;

                await itemExacto.WaitForAsync(new() { Timeout = 60000 });
                await itemExacto.DblClickAsync(new() { Force = true });

                await page.WaitForFunctionAsync(
                    @"(p) => {
                        const sel = document.querySelector('#lstFiltros');
                        if (!sel) return false;
                        return Array.from(sel.options || []).some(o => (o.textContent||'').includes(p));
                    }",
                    partida,
                    new() { Timeout = 60000 }
                );

                var prev = DateTime.Today.AddMonths(-1);
                string[] meses = { "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic" };
                var mesUi = meses[prev.Month - 1];
                var esperado = $"{mesUi} {prev:yyyy}";

                async Task PickPrevMonthAsync(ILocator input)
                {
                await input.WaitForAsync(new() { Timeout = 60000 });
                await input.ClickAsync(new() { Force = true });

                var picker = page.Locator(".datepicker:visible, .datepicker-dropdown:visible").First;
                    try
                    {
                await picker.WaitForAsync(new() { Timeout = 1500 });
                    }
                    catch
                    {
                var iconBtn = input.Locator("xpath=ancestor::*[contains(@class,'input-group')][1]//button").First;
                        if (await iconBtn.CountAsync() == 0)
                            iconBtn = input.Locator("xpath=ancestor::*[contains(@class,'input-group')][1]//span[contains(@class,'input-group-addon') or contains(@class,'add-on')]").First;

                        if (await iconBtn.CountAsync() > 0)
                            await iconBtn.ClickAsync(new() { Force = true });
                        else
                            await input.Locator("xpath=following-sibling::*[1]").ClickAsync(new() { Force = true });

                        picker = page.Locator(".datepicker:visible, .datepicker-dropdown:visible").First;
                        await picker.WaitForAsync(new() { Timeout = 60000 });
                    }

                var monthsView = picker.Locator(".datepicker-months");
                    if (await monthsView.CountAsync() > 0 && !await monthsView.IsVisibleAsync())
                    {
                        var sw = picker.Locator("th.datepicker-switch").First;
                        await sw.ClickAsync(new() { Force = true });
                        if (!await monthsView.IsVisibleAsync())
                            await sw.ClickAsync(new() { Force = true });
                    }

                    await monthsView.WaitForAsync(new() { Timeout = 60000 });

                var mesBtn = picker.Locator($".datepicker-months span.month:not(.disabled)", new() { HasTextString = mesUi }).First;
                await mesBtn.WaitForAsync(new() { Timeout = 60000 });
                await mesBtn.ClickAsync(new() { Force = true });

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

                var fromInput = page.Locator("#cboDesde input").First;
                await PickPrevMonthAsync(fromInput);

                var toInput = page.Locator("#cboHasta input").First;
                await PickPrevMonthAsync(toInput);

                Console.WriteLine($"Desde={(await fromInput.InputValueAsync()).Trim()} | Hasta={(await toInput.InputValueAsync()).Trim()} | Esperado={esperado}");

                await page.Locator("#btnBuscar").ClickAsync(new() { Force = true });
                await page.WaitForTimeoutAsync(800);

                var noResultsOk = page.Locator("#btnOKModalVentanaMensaje");
                if (await noResultsOk.CountAsync() > 0 && await noResultsOk.IsVisibleAsync())
                {
                    Console.WriteLine($"[{partida}] No se encontraron registros. Aceptar + Restablecer + siguiente código.");

                await noResultsOk.ClickAsync(new() { Force = true });
                await page.WaitForTimeoutAsync(800);

                await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });
                await page.WaitForTimeoutAsync(1200);

                await page.WaitForFunctionAsync(
                    @"() => {
                        const el = document.querySelector('#txtNandinaB');
                        return el && !el.disabled;
                    }",
                    null,
                    new() { Timeout = 30000 }

                    );

                    continue;
                }

                var start = DateTime.UtcNow;
                while (true)
                {
                    if (await noResultsOk.CountAsync() > 0 && await noResultsOk.IsVisibleAsync())
                    {
                        Console.WriteLine($"[{partida}] No se encontraron registros (tardío). Aceptar + Restablecer + siguiente código.");

                        await noResultsOk.ClickAsync(new() { Force = true });
                        await page.WaitForTimeoutAsync(800);

                        await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });
                        await page.WaitForTimeoutAsync(1200);

                        await page.WaitForFunctionAsync(
                        @"() => {
                            const el = document.querySelector('#txtNandinaB');
                            return el && !el.disabled;
                        }",
                        null,
                        new() { Timeout = 30000 }
                    );

                        continue;
                    }

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

                    if ((DateTime.UtcNow - start).TotalMilliseconds > 180000)
                        throw new TimeoutException($"[{partida}] Timeout esperando '#totalRecordsFound' con 'Se encontraron ... registros'.");

                    await page.WaitForTimeoutAsync(500);
                }

                var desdeFiltroPre = (await fromInput.InputValueAsync()).Trim();
                var hastaFiltroPre = (await toInput.InputValueAsync()).Trim();
                var tipoKey = "exportacion";
                var excelKeyPre = $"{Paises.Peru}|{tipoKey}|{partida}|{desdeFiltroPre}|{hastaFiltroPre}".Trim();

                if (await repo.ExcelKeyExistsAsync(excelKeyPre))
                {
                    Console.WriteLine($"[SKIP-PRE] Ya existe ExcelKey={excelKeyPre}. No se descarga para evitar duplicados/ban.");
                    await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });
                    await page.WaitForTimeoutAsync(1200);
                    continue;
                }

                    string fullPath = "";
                for (int attemptDl = 1; attemptDl <= 2; attemptDl++)
                {
                    try
                    {
                        await page.Locator("#tabDetalleExcel").ClickAsync(new() { Force = true });

                        var excelBtn = page.Locator("#downloadFileVerRegistro2");
                        await excelBtn.WaitForAsync(new() { Timeout = 120000, State = WaitForSelectorState.Attached });

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

                        var downloadTask = page.WaitForDownloadAsync(new() { Timeout = 600000 });
                        await excelBtn.ClickAsync(new() { Force = true });
                        var download = await downloadTask;

                        var filename = download.SuggestedFilename;
                        if (string.IsNullOrWhiteSpace(filename))
                            filename = "veritrade_detalle.xlsx";

                        fullPath = Path.Combine(downloadsDir, $"{DateTime.Now:yyyyMMdd_HHmmss}_{partida}_{filename}");
                        await download.SaveAsAsync(fullPath);

                        Console.WriteLine($"Excel descargado: {fullPath}");
                        break; 
                    }
                    catch (TimeoutException tex)
                    {
                        Console.WriteLine($"[WARN] Timeout descargando Excel (attempt {attemptDl}/2) partida={partida}: {tex.Message}");

                 
                        try
                        {
                            await page.GotoAsync("https://www.veritradecorp.com/es/mis-busquedas", new() { Timeout = 120000, WaitUntil = WaitUntilState.DOMContentLoaded });
                            await page.WaitForSelectorAsync("#btnBuscar", new() { Timeout = 120000 });
                        }
                        catch { }

                        if (attemptDl == 2) throw;
                        await page.WaitForTimeoutAsync(8000);
                    }
                }
                if (string.IsNullOrWhiteSpace(fullPath))
                    Console.WriteLine($"Excel descargado: {fullPath}");

                var rows = reader.Read(fullPath, Paises.Peru);

                //validacion duplicidad
                var desdeFiltro = (await fromInput.InputValueAsync()).Trim();
                var hastaFiltro = (await toInput.InputValueAsync()).Trim();

                // tipo ya viene dentro del Excel (tu reader lo detecta), pero para el key necesitamos el mismo string
                // lo más práctico: usar el del primer row luego de leer
                var tipo = rows.Count > 0 ? rows[0].Tipo : "";

                var excelKey = $"{Paises.Peru}|{tipo}|{partida}|{desdeFiltro}|{hastaFiltro}".Trim();

                // “inyectar” los 3 campos a todas las filas sin romper el modelo
                rows = rows.Select(r => new VeritradeRow
                {
                    PaisCarga = r.PaisCarga,
                    Tipo = r.Tipo,
                    PartidaAduanera = r.PartidaAduanera,
                    DescripcionPartidaAduanera = r.DescripcionPartidaAduanera,
                    Aduana = r.Aduana,
                    DUA_DAM = r.DUA_DAM,
                    Fecha = r.Fecha,
                    ETA = r.ETA,
                    ManifiestoNr = r.ManifiestoNr,
                    CodTributario = r.CodTributario,
                    Importador = r.Importador,
                    Exportador = r.Exportador,
                    EmbarcadorExportador = r.EmbarcadorExportador,
                    PaisdeCompra = r.PaisdeCompra,
                    PuertodeEmbarque = r.PuertodeEmbarque,
                    FechadeEmbarque = r.FechadeEmbarque,
                    Marca = r.Marca,
                    PaisdeEmbarque = r.PaisdeEmbarque,
                    Producto = r.Producto,
                    PaisdelExportador = r.PaisdelExportador,
                    RegAduana1 = r.RegAduana1,
                    KgBruto = r.KgBruto,
                    KgNeto = r.KgNeto,
                    Qty1 = r.Qty1,
                    Und1 = r.Und1,
                    Qty2 = r.Qty2,
                    Und2 = r.Und2,
                    US_FOB_Tot = r.US_FOB_Tot,
                    US_CFR_Tot = r.US_CFR_Tot,
                    US_CIF_Tot = r.US_CIF_Tot,
                    PaisOrigen = r.PaisOrigen,
                    Via = r.Via,
                    DescripcionComercial = r.DescripcionComercial,
                    Descripcion1 = r.Descripcion1,
                    Descripcion2 = r.Descripcion2,
                    Descripcion3 = r.Descripcion3,
                    Descripcion4 = r.Descripcion4,
                    Descripcion5 = r.Descripcion5,

                    DesdeFiltro = desdeFiltro,
                    HastaFiltro = hastaFiltro,
                    ExcelKey = excelKey
                }).ToList();


                Console.WriteLine($"Excel leído: {Path.GetFileName(fullPath)} | rows.Count={rows.Count}");

                var cargaId = Guid.NewGuid().ToString("N");
                var cargaFecha = DateTime.UtcNow;

                Console.WriteLine($"Insertando en SQL... cargaId={cargaId} cargaFecha={cargaFecha:O}");
                await repo.InsertAsync(rows, cargaId, cargaFecha);

                Console.WriteLine($"Insertado SQL: filas(leídas)={rows.Count}, cargaId={cargaId}");

                await page.Locator("#btnRestablecer").ClickAsync(new() { Force = true });
                await page.WaitForTimeoutAsync(1200);
                await page.WaitForFunctionAsync(
                    @"() => {
                            const el = document.querySelector('#txtNandinaB');
                            return el && !el.disabled;
                        }",
                    null,
                    new() { Timeout = 30000 }
                );
            }
        }
    }
}
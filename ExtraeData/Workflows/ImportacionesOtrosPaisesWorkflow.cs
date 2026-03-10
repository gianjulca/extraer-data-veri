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
    public static class ImportacionesOtrosPaisesWorkflow
    {
        public static async Task RunAsync(
            IPage page,
            string downloadsDir,
            VeritradeExcelReader reader,
            VeritradeSqlRepository repo)
        {
            async Task SelectCountryAsync(string label)
            {
                var cbo = page.Locator("#cboPais");
                if (await cbo.CountAsync() == 0)
                    cbo = page.Locator("#cboPais2");

                await cbo.WaitForAsync(new() { Timeout = 60000 });

                var handle = await cbo.ElementHandleAsync();

                await page.WaitForFunctionAsync(
                    @"(sel) => sel && sel.options && sel.options.length > 1",
                    handle,
                    new() { Timeout = 60000 }
                );

                var ok = await page.EvaluateAsync<bool>(
                    @"(p) => {
                        const sel = p.sel;
                        const txt = p.txt;
                        const norm = s => (s || '').trim().toLowerCase();
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

            async Task WaitTxtNandinaReadyAsync()
            {
                Exception? lastEx = null;

                for (int attempt = 1; attempt <= 3; attempt++)
                {
                    try
                    {
                        var partidaInputReady = page.Locator("#txtNandinaB");
                        await partidaInputReady.WaitForAsync(new() { Timeout = 10000 });

                        await page.WaitForFunctionAsync(
                            @"() => {
                                const el = document.querySelector('#txtNandinaB');
                                const loader = document.querySelector('#loadingPageAdmin');

                                const loaderActive =
                                    loader &&
                                    loader.classList.contains('is-active-loading') &&
                                    window.getComputedStyle(loader).display !== 'none' &&
                                    window.getComputedStyle(loader).visibility !== 'hidden';

                                return el && !el.disabled && !loaderActive;
                            }",
                            null,
                            new() { Timeout = 2000 }
                        );

                        return;
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                        Console.WriteLine($"[TXT-NANDINA-WARN-OTROS] intento {attempt}/3 esperando #txtNandinaB habilitado: {ex.Message}");
                        await page.WaitForTimeoutAsync(1000);
                    }
                }

                throw new Exception($"No se pudo dejar habilitado #txtNandinaB. Último error: {lastEx?.Message}");
            }

            async Task ResetAndWaitReadyAsync()
            {
                Exception? lastEx = null;

                for (int attempt = 1; attempt <= 3; attempt++)
                {
                    try
                    {
                        var resetBtn = page.Locator("#btnRestablecer");
                        await resetBtn.WaitForAsync(new() { Timeout = 10000 });
                        await resetBtn.ClickAsync(new() { Force = true, Timeout = 5000 });
                        await page.WaitForTimeoutAsync(1500);

                        await WaitTxtNandinaReadyAsync();
                        return;
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                        Console.WriteLine($"[RESET-WARN-OTROS] intento {attempt}/3 falló: {ex.Message}");
                        await page.WaitForTimeoutAsync(1500);
                    }
                }

                throw new Exception($"No se pudo restablecer filtros y habilitar #txtNandinaB. Último error: {lastEx?.Message}");
            }

            async Task TypePartidaAsync(ILocator partidaInput, string partida)
            {
                Exception? lastEx = null;

                for (int attempt = 1; attempt <= 4; attempt++)
                {
                    try
                    {
                        await WaitTxtNandinaReadyAsync();

                        await partidaInput.ClickAsync(new() { Timeout = 5000 });
                        await partidaInput.FillAsync("", new() { Timeout = 5000 });
                        await page.WaitForTimeoutAsync(250);

                        await partidaInput.TypeAsync(partida, new() { Delay = 80, Timeout = 5000 });
                        await page.WaitForTimeoutAsync(800);

                        var current = await partidaInput.InputValueAsync();

                        if (!string.IsNullOrWhiteSpace(current) && current.Contains(partida))
                        {
                            var suggestions = page.Locator("li.ui-menu-item:visible");

                            try
                            {
                                await suggestions.First.WaitForAsync(new() { Timeout = 2000 });
                            }
                            catch
                            {
                            }

                            if (await suggestions.CountAsync() > 0)
                                return;
                        }

                        lastEx = new Exception($"Autocomplete no visible para '{partida}' en intento {attempt}/4.");
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                    }

                    Console.WriteLine($"[TYPE-WARN-OTROS] intento {attempt}/4 para partida={partida} falló: {lastEx?.Message}");

                    if (attempt < 4)
                    {
                        try
                        {
                            await ResetAndWaitReadyAsync();
                        }
                        catch
                        {
                            await page.WaitForTimeoutAsync(1000);
                        }
                    }
                }

                throw new Exception($"No se pudo escribir/mostrar autocomplete para '{partida}' en #txtNandinaB. Último error: {lastEx?.Message}");
            }

            async Task PickPrevMonthAsync(ILocator input, string esperado, string mesUi)
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


            async Task RunForCountryAsync(string countryLabel, string[] countryPartidas)
            {
                if (countryPartidas == null || countryPartidas.Length == 0)
                {
                    Console.WriteLine($"⚠️ {countryLabel}: no hay partidas definidas, se omite.");
                    return;
                }

                Console.WriteLine($"\n===== INICIANDO PAÍS: {countryLabel} | partidas={countryPartidas.Length} =====");

                var rdbImp = page.Locator("#rdbImp");
                if (await rdbImp.CountAsync() > 0 && !await rdbImp.IsCheckedAsync())
                    await rdbImp.CheckAsync(new() { Force = true });

                await SelectCountryAsync(countryLabel);

                for (int j = 0; j < countryPartidas.Length; j++)
                {
                    var partida = countryPartidas[j];

                    for (int retryPartida = 1; retryPartida <= 2; retryPartida++)
                    {
                        try
                        {
                            Console.WriteLine($"\n=== [{countryLabel}] Partida {partida} ({j + 1}/{countryPartidas.Length}) | intento {retryPartida}/2 ===");

                            var partidaInput = page.Locator("#txtNandinaB");
                            await partidaInput.WaitForAsync(new() { Timeout = 10000 });

                            await WaitTxtNandinaReadyAsync();

                            await TypePartidaAsync(partidaInput, partida);

                            var itemExacto = page.Locator($"li.ui-menu-item:visible >> text={partida}").First;
                            if (await itemExacto.CountAsync() == 0)
                                itemExacto = page.Locator("li.ui-menu-item:visible").First;

                            await itemExacto.WaitForAsync(new() { Timeout = 60000 });
                            await itemExacto.DblClickAsync(new() { Force = true, Timeout = 5000 });

                            await page.WaitForFunctionAsync(
                                @"(p) => {
                                    const sel = document.querySelector('#lstFiltros');
                                    if (!sel) return false;
                                    return Array.from(sel.options || []).some(o => (o.textContent || '').includes(p));
                                }",
                                partida,
                                new() { Timeout = 60000 }
                            );



                            var fromInput = page.Locator("#cboDesde input").First;
                            var toInput = page.Locator("#cboHasta input").First;

                            if (countryLabel == Paises.Ecuador)
                            {
                                var prev = DateTime.Today.AddMonths(-1);
                                string[] meses = { "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic" };
                                var mesUi = meses[prev.Month - 1];
                                var esperado = $"{mesUi} {prev:yyyy}";

                                await PickPrevMonthAsync(fromInput, esperado, mesUi);
                                await PickPrevMonthAsync(toInput, esperado, mesUi);

                                Console.WriteLine($"[FECHAS ECUADOR] Desde={(await fromInput.InputValueAsync()).Trim()} | Hasta={(await toInput.InputValueAsync()).Trim()} | Esperado={esperado}");
                            }
                            else
                            {
                                Console.WriteLine($"[FECHAS DEFAULT {countryLabel}] Desde={(await fromInput.InputValueAsync()).Trim()} | Hasta={(await toInput.InputValueAsync()).Trim()}");
                            }





                            await page.Locator("#btnBuscar").ClickAsync(new() { Force = true });
                            await page.WaitForTimeoutAsync(800);

                            var noResultsOk = page.Locator("#btnOKModalVentanaMensaje");
                            if (await noResultsOk.CountAsync() > 0 && await noResultsOk.IsVisibleAsync())
                            {
                                Console.WriteLine($"[{partida}] No se encontraron registros. Aceptar + Restablecer + siguiente código.");

                                await noResultsOk.ClickAsync(new() { Force = true });
                                await page.WaitForTimeoutAsync(800);

                                await ResetAndWaitReadyAsync();
                                break;
                            }

                            var start = DateTime.UtcNow;
                            bool foundResults = false;

                            while (true)
                            {
                                if (await noResultsOk.CountAsync() > 0 && await noResultsOk.IsVisibleAsync())
                                {
                                    Console.WriteLine($"[{partida}] No se encontraron registros (tardío). Aceptar + Restablecer + siguiente código.");

                                    await noResultsOk.ClickAsync(new() { Force = true });
                                    await page.WaitForTimeoutAsync(800);

                                    await ResetAndWaitReadyAsync();
                                    break;
                                }

                                var total = page.Locator("#totalRecordsFound");
                                if (await total.CountAsync() > 0)
                                {
                                    var t = (await total.InnerTextAsync()).Trim();
                                    if (t.IndexOf("Se encontraron", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                        t.IndexOf("registros", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        Console.WriteLine($"[{partida}] OK => {t}");
                                        foundResults = true;
                                        break;
                                    }
                                }

                                if ((DateTime.UtcNow - start).TotalMilliseconds > 180000)
                                    throw new TimeoutException($"[{partida}] Timeout esperando '#totalRecordsFound' con 'Se encontraron ... registros'.");

                                await page.WaitForTimeoutAsync(800);
                            }

                            if (!foundResults)
                                break;

                            var desdeFiltroPre = (await fromInput.InputValueAsync()).Trim();
                            var hastaFiltroPre = (await toInput.InputValueAsync()).Trim();
                            var tipoKey = "importacion";
                            var excelKeyPre = $"{countryLabel}|{tipoKey}|{partida}|{desdeFiltroPre}|{hastaFiltroPre}".Trim();

                            if (await repo.ExcelKeyExistsAsync(excelKeyPre))
                            {
                                Console.WriteLine($"[SKIP-PRE] Ya existe ExcelKey={excelKeyPre}. No se descarga para evitar duplicados/ban.");
                                await ResetAndWaitReadyAsync();
                                break;
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

                                    var downloadTask = page.WaitForDownloadAsync(new() { Timeout = 60000 });
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
                                        await ResetAndWaitReadyAsync();
                                        await SelectCountryAsync(countryLabel);
                                        await page.WaitForTimeoutAsync(800);
                                    }
                                    catch
                                    {
                                    }

                                    if (attemptDl == 2)
                                        throw;

                                    await page.WaitForTimeoutAsync(2000);
                                }
                            }

                            if (string.IsNullOrWhiteSpace(fullPath))
                                throw new Exception($"No se pudo descargar el Excel para partida={partida} (fullPath vacío).");

                            Console.WriteLine($"Excel descargado: {fullPath}");

                            var rows = reader.Read(fullPath, countryLabel);

                            var desdeFiltro = (await fromInput.InputValueAsync()).Trim();
                            var hastaFiltro = (await toInput.InputValueAsync()).Trim();
                            var tipo = rows.Count > 0 ? rows[0].Tipo : "";
                            var excelKey = $"{countryLabel}|{tipo}|{partida}|{desdeFiltro}|{hastaFiltro}".Trim();

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

                            foreach (var r in rows.Take(3))
                                Console.WriteLine($"[HEAD] {r.PartidaAduanera} | {r.DescripcionPartidaAduanera} | {r.Importador}");

                            foreach (var r in rows.TakeLast(3))
                                Console.WriteLine($"[TAIL] {r.PartidaAduanera} | {r.DescripcionPartidaAduanera} | {r.Importador}");

                            var vacias = rows.Count(r =>
                                string.IsNullOrWhiteSpace(r.PartidaAduanera) &&
                                string.IsNullOrWhiteSpace(r.DescripcionPartidaAduanera) &&
                                string.IsNullOrWhiteSpace(r.Importador));

                            Console.WriteLine($"Filas totalmente vacías (3 campos): {vacias}");

                            var cargaId = Guid.NewGuid().ToString("N");
                            var cargaFecha = DateTime.UtcNow;

                            Console.WriteLine($"Insertando en SQL... cargaId={cargaId} cargaFecha={cargaFecha:O}");
                            await repo.InsertAsync(rows, cargaId, cargaFecha);

                            Console.WriteLine($"Insertado SQL: filas(leídas)={rows.Count}, cargaId={cargaId}");

                            await ResetAndWaitReadyAsync();
                            break;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[RETRY-OTROS] {countryLabel} partida={partida} intento {retryPartida}/2 falló: {ex.Message}");

                            if (retryPartida == 2)
                                throw;

                            try
                            {
                                await ResetAndWaitReadyAsync();
                            }
                            catch
                            {
                            }

                            try
                            {
                                await SelectCountryAsync(countryLabel);
                                await page.WaitForTimeoutAsync(800);
                            }
                            catch
                            {
                            }

                            await page.WaitForTimeoutAsync(1500);
                        }
                    }
                }
            }

            await RunForCountryAsync(Paises.Chile, Partidas.Chile);
            await RunForCountryAsync(Paises.Colombia, Partidas.Colombia);
            await RunForCountryAsync(Paises.Ecuador, Partidas.Ecuador);
        }
    }
}



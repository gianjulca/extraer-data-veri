/*using ClosedXML.Excel;
using ExtraeData.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtraeData.Services
{
    public sealed class VeritradeExcelReader
    {
        public List<VeritradeRow> Read(string filePath)
        {
            using var wb = new XLWorkbook(filePath);
            var ws = wb.Worksheets.First();

            const string colPartida = "Partida Aduanera";
            const string colDesc = "Descripcion de la Partida Aduanera";
            const string colImportador = "Importador";

            var headerRow = ws.RowsUsed()
                .FirstOrDefault(r => r.CellsUsed().Any(c =>
                    string.Equals(c.GetString().Trim(), colPartida, StringComparison.OrdinalIgnoreCase)));

            if (headerRow is null)
                throw new Exception($"No se encontró la fila de encabezados con '{colPartida}'.");

            int headerRowNum = headerRow.RowNumber();

            int cPartida = headerRow.CellsUsed()
      .First(c => string.Equals(c.GetString().Trim(), colPartida, StringComparison.OrdinalIgnoreCase))
      .Address.ColumnNumber;

            int cDesc = headerRow.CellsUsed()
                .First(c => string.Equals(c.GetString().Trim(), colDesc, StringComparison.OrdinalIgnoreCase))
                .Address.ColumnNumber;

            int cImportador = headerRow.CellsUsed()
                .First(c => string.Equals(c.GetString().Trim(), colImportador, StringComparison.OrdinalIgnoreCase))
                .Address.ColumnNumber;

            int lastRow = ws.LastRowUsed().RowNumber();

            var rows = new List<VeritradeRow>();

            for (int r = headerRowNum + 1; r <= lastRow; r++)
            {
                var partida = ws.Cell(r, cPartida).GetString().Trim();
                var desc = ws.Cell(r, cDesc).GetString().Trim();
                var importador = ws.Cell(r, cImportador).GetString().Trim();

                if (string.IsNullOrWhiteSpace(partida) && string.IsNullOrWhiteSpace(importador) && string.IsNullOrWhiteSpace(desc))
                    continue;

                rows.Add(new VeritradeRow
                {
                    PartidaAduanera = partida,
                    DescripcionPartidaAduanera = desc,
                    Importador = importador
                });
            }

            return rows;
        }

    }
}*/

using ClosedXML.Excel;
using ExtraeData.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExtraeData.Services
{
    public sealed class VeritradeExcelReader
    {

        private static int? FindCol(IXLRow headerRow, params string[] names)
        {
            foreach (var name in names)
            {
                var cell = headerRow.CellsUsed()
                    .FirstOrDefault(c => string.Equals(c.GetString().Trim(), name, StringComparison.OrdinalIgnoreCase));

                if (cell != null)
                    return cell.Address.ColumnNumber;
            }
            return null;
        }

        private static DateTime? ReadDate(IXLWorksheet ws, int row, int col)
        {
            var c = ws.Cell(row, col);
            if (c.IsEmpty()) return null;

            if (c.DataType == XLDataType.DateTime)
                return c.GetDateTime().Date;

            var s = c.GetString().Trim();
            if (DateTime.TryParse(s, out var dt))
                return dt.Date;

            return null;
        }

        private static decimal? ReadDecimal(IXLWorksheet ws, int row, int col)
        {
            var c = ws.Cell(row, col);
            if (c.IsEmpty()) return null;

            if (c.DataType == XLDataType.Number)
                return (decimal)c.GetDouble();

            var s = c.GetString().Trim();
            if (string.IsNullOrWhiteSpace(s)) return null;

            // normaliza separadores comunes
            s = s.Replace(",", "");
            if (decimal.TryParse(s, out var d))
                return d;

            return null;
        }

        public List<VeritradeRow> Read(string filePath, string paisCarga)
        {
            using var wb = new XLWorkbook(filePath);
            var ws = wb.Worksheets.First();

            //Detectar TIPO desde el título (PERU - IMPORTACIONES / PERU - EXPORTACIONES)
            var headText = string.Join(" ", ws.Range("A1:G5").Cells().Select(c => c.GetString())).ToUpperInvariant();
            var tipo = headText.Contains("EXPORTACIONES") ? "exportacion"
                     : headText.Contains("IMPORTACIONES") ? "importacion"
                     : "";

            const string colPartida = "Partida Aduanera";
            const string colDesc = "Descripcion de la Partida Aduanera";          

            const string colAduana = "Aduana";
            const string colDuaDam = "DUA / DAM";
            const string colFecha = "Fecha";
            const string colEta = "ETA";
            const string colManifiesto = "Manifiesto Nr.";
            const string colCodTrib = "Cod. Tributario";

            const string colImportador = "Importador";
            const string colEmbarExportador = "Embarcador / Exportador";
            const string colExportador = "Exportador";                 

            const string colPaisdeCompra = "Pais de Compra";
            const string colPuertodeEmbarque = "Puerto de Embarque";
            const string colFechadeEmbarque = "Fecha de Embarque";
            const string colMarca = "Marca";
            const string colPaisdeEmbarque = "Pais de Embarque";
            const string colProducto = "Producto";
            const string colExportadorReal = "Exportador";
            const string colPaisdelExportador = "Pais del Exportador";
            const string colRegAduana1 = "Reg Aduana 1";


            const string colKgBruto = "Kg Bruto";
            const string colKgNeto = "Kg Neto";
            const string colQty1 = "Qty 1";
            const string colUnd1 = "Und 1";
            const string colQty2 = "Qty 2";
            const string colUnd2 = "Und 2";

            const string colFob = "U$ FOB Tot";
            const string colCfr = "U$ CFR Tot";
            const string colCif = "U$ CIF Tot";

            const string colPaisOrigen = "Pais de Origen";
            const string colVia = "Via";
            const string colDescCom = "Descripcion Comercial";
            const string colDesc1 = "Descripcion1";
            const string colDesc2 = "Descripcion2";
            const string colDesc3 = "Descripcion3";
            const string colDesc4 = "Descripcion4";
            const string colDesc5 = "Descripcion5";

            var headerRow = ws.RowsUsed()
                .FirstOrDefault(r => r.CellsUsed().Any(c =>
                    string.Equals(c.GetString().Trim(), colPartida, StringComparison.OrdinalIgnoreCase)));

            if (headerRow is null)
                throw new Exception($"No se encontró la fila de encabezados con '{colPartida}'.");

            int headerRowNum = headerRow.RowNumber();

            int cPartida = FindCol(headerRow, colPartida) ?? throw new Exception($"No se encontró columna '{colPartida}'.");
            int cDesc = FindCol(headerRow, colDesc) ?? throw new Exception($"No se encontró columna '{colDesc}'.");
            int? cImportador = FindCol(headerRow, colImportador); // puede no existir en exportaciones

            int? cAduana = FindCol(headerRow, colAduana);
            int? cDuaDam = FindCol(headerRow, colDuaDam);
            int? cFecha = FindCol(headerRow, colFecha);
            int? cEta = FindCol(headerRow, colEta);
            int? cManifiesto = FindCol(headerRow, colManifiesto);
            int? cCodTrib = FindCol(headerRow, colCodTrib);
            int? cEmbExp = FindCol(headerRow, colEmbarExportador); // "Embarcador / Exportador"     

            int? cPaisdeCompra = FindCol(headerRow, colPaisdeCompra);
            int? cPuertodeEmbarque = FindCol(headerRow, colPuertodeEmbarque);
            int? cFechadeEmbarque = FindCol(headerRow, colFechadeEmbarque);
            int? cMarca = FindCol(headerRow, colMarca);
            int? cPaisdeEmbarque = FindCol(headerRow, colPaisdeEmbarque);
            int? cProducto = FindCol(headerRow, colProducto);
            int? cPaisdelExportador = FindCol(headerRow, colPaisdelExportador);
            int? cRegAduana1 = FindCol(headerRow, colRegAduana1);
            int? cExportador = FindCol(headerRow, colExportador); // "Exportador"


            int? cKgBruto = FindCol(headerRow, colKgBruto);
            int? cKgNeto = FindCol(headerRow, colKgNeto);
            int? cQty1 = FindCol(headerRow, colQty1);
            int? cUnd1 = FindCol(headerRow, colUnd1);
            int? cQty2 = FindCol(headerRow, colQty2);
            int? cUnd2 = FindCol(headerRow, colUnd2);

            int? cFob = FindCol(headerRow, colFob);
            int? cCfr = FindCol(headerRow, colCfr);
            int? cCif = FindCol(headerRow, colCif);

            int? cPaisOrigen = FindCol(headerRow, colPaisOrigen);
            int? cVia = FindCol(headerRow, colVia);
            int? cDescCom = FindCol(headerRow, colDescCom);
            int? cDesc1 = FindCol(headerRow, colDesc1);
            int? cDesc2 = FindCol(headerRow, colDesc2);
            int? cDesc3 = FindCol(headerRow, colDesc3);
            int? cDesc4 = FindCol(headerRow, colDesc4);
            int? cDesc5 = FindCol(headerRow, colDesc5);

            var rows = new List<VeritradeRow>();

            var last = ws.LastRowUsed();
            if (last == null) return rows;
            int lastRow = last.RowNumber();

            for (int r = headerRowNum + 1; r <= lastRow; r++)
            {
                var partida = ws.Cell(r, cPartida).GetString().Trim();
                var desc = ws.Cell(r, cDesc).GetString().Trim();

                var importador = (cImportador.HasValue ? ws.Cell(r, cImportador.Value).GetString().Trim() : "");

                var embarcadorExportador = (cEmbExp.HasValue ? ws.Cell(r, cEmbExp.Value).GetString().Trim() : "");
                var exportador = (cExportador.HasValue ? ws.Cell(r, cExportador.Value).GetString().Trim() : "");

                if (string.IsNullOrWhiteSpace(partida) && string.IsNullOrWhiteSpace(desc) &&
                         string.IsNullOrWhiteSpace(importador) &&
                         string.IsNullOrWhiteSpace(embarcadorExportador) &&
                         string.IsNullOrWhiteSpace(exportador))
                         continue;

                rows.Add(new VeritradeRow
                {
                    Tipo = tipo,
                    PaisCarga = paisCarga,

                    PartidaAduanera = partida,
                    DescripcionPartidaAduanera = desc,
                    Aduana = cAduana.HasValue ? ws.Cell(r, cAduana.Value).GetString().Trim() : "",
                    DUA_DAM = cDuaDam.HasValue ? ws.Cell(r, cDuaDam.Value).GetString().Trim() : "",

                    Fecha = cFecha.HasValue ? ReadDate(ws, r, cFecha.Value) : null,
                    ETA = cEta.HasValue ? ReadDate(ws, r, cEta.Value) : null,

                    ManifiestoNr = cManifiesto.HasValue ? ws.Cell(r, cManifiesto.Value).GetString().Trim() : "",
                    CodTributario = cCodTrib.HasValue ? ws.Cell(r, cCodTrib.Value).GetString().Trim() : "",

                    Importador = importador,
                    EmbarcadorExportador = embarcadorExportador,
                    Exportador = exportador,

                    PaisdeCompra = cPaisdeCompra.HasValue ? ws.Cell(r, cPaisdeCompra.Value).GetString().Trim() : "",
                    PuertodeEmbarque = cPuertodeEmbarque.HasValue ? ws.Cell(r, cPuertodeEmbarque.Value).GetString().Trim() : "",
                    FechadeEmbarque = cFechadeEmbarque.HasValue ? ReadDate(ws, r, cFechadeEmbarque.Value) : null,
                    Marca = cMarca.HasValue ? ws.Cell(r, cMarca.Value).GetString().Trim() : "",
                    PaisdeEmbarque = cPaisdeEmbarque.HasValue ? ws.Cell(r, cPaisdeEmbarque.Value).GetString().Trim() : "",
                    Producto = cProducto.HasValue ? ws.Cell(r, cProducto.Value).GetString().Trim() : "",
                    PaisdelExportador = cPaisdelExportador.HasValue ? ws.Cell(r, cPaisdelExportador.Value).GetString().Trim() : "",
                    RegAduana1 = cRegAduana1.HasValue ? ws.Cell(r, cRegAduana1.Value).GetString().Trim() : "",


                    KgBruto = cKgBruto.HasValue ? ReadDecimal(ws, r, cKgBruto.Value) : null,
                    KgNeto = cKgNeto.HasValue ? ReadDecimal(ws, r, cKgNeto.Value) : null,

                    Qty1 = cQty1.HasValue ? ReadDecimal(ws, r, cQty1.Value) : null,
                    Und1 = cUnd1.HasValue ? ws.Cell(r, cUnd1.Value).GetString().Trim() : "",

                    Qty2 = cQty2.HasValue ? ReadDecimal(ws, r, cQty2.Value) : null,
                    Und2 = cUnd2.HasValue ? ws.Cell(r, cUnd2.Value).GetString().Trim() : "",

                    US_FOB_Tot = cFob.HasValue ? ReadDecimal(ws, r, cFob.Value) : null,
                    US_CFR_Tot = cCfr.HasValue ? ReadDecimal(ws, r, cCfr.Value) : null,
                    US_CIF_Tot = cCif.HasValue ? ReadDecimal(ws, r, cCif.Value) : null,

                    PaisOrigen = cPaisOrigen.HasValue ? ws.Cell(r, cPaisOrigen.Value).GetString().Trim() : "",
                    Via = cVia.HasValue ? ws.Cell(r, cVia.Value).GetString().Trim() : "",

                    DescripcionComercial = cDescCom.HasValue ? ws.Cell(r, cDescCom.Value).GetString().Trim() : "",
                    Descripcion1 = cDesc1.HasValue ? ws.Cell(r, cDesc1.Value).GetString().Trim() : "",
                    Descripcion2 = cDesc2.HasValue ? ws.Cell(r, cDesc2.Value).GetString().Trim() : "",
                    Descripcion3 = cDesc3.HasValue ? ws.Cell(r, cDesc3.Value).GetString().Trim() : "",
                    Descripcion4 = cDesc4.HasValue ? ws.Cell(r, cDesc4.Value).GetString().Trim() : "",
                    Descripcion5 = cDesc5.HasValue ? ws.Cell(r, cDesc5.Value).GetString().Trim() : "",
                });
            }

            return rows;
        }
    }
}

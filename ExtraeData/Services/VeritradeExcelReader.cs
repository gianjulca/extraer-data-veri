using ClosedXML.Excel;
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
}

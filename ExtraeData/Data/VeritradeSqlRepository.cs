/*using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExtraeData.Models;
using Microsoft.Data.SqlClient;

namespace ExtraeData.Data
{
    public sealed class VeritradeSqlRepository
    {
        private readonly string _cs;

        public VeritradeSqlRepository(string connectionString)
        {
            _cs = connectionString;
        }

        public async Task InsertAsync(IEnumerable<VeritradeRow> rows, String cargaId, DateTime cargaFecha)
        {
            const string sql = @"
INSERT INTO dbo.ImportacionesAduanas
(PartidaAduanera, DescripcionPartidaAduanera, Importador,CargaId, CargaFecha)
VALUES
(@Partida, @Desc, @Importador, @CargaId, @CargaFecha);";

            using var con = new SqlConnection(_cs);
            await con.OpenAsync();

            foreach (var r in rows)
            {
                using var cmd = new SqlCommand(sql, con);
                cmd.Parameters.AddWithValue("@Partida", (object?)r.PartidaAduanera ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Desc", (object?)r.DescripcionPartidaAduanera ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Importador", (object?)r.Importador ?? DBNull.Value);

                cmd.Parameters.AddWithValue("@CargaId", cargaId);
                cmd.Parameters.AddWithValue("@CargaFecha", cargaFecha);

                await cmd.ExecuteNonQueryAsync();
            }
        }
    }
}
*/

using DocumentFormat.OpenXml.Spreadsheet;
using ExtraeData.Models;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtraeData.Data
{
    public sealed class VeritradeSqlRepository
    {
        private readonly string _cs;

        public VeritradeSqlRepository(string connectionString)
        {
            _cs = connectionString;
        }

        public async Task InsertAsync(IEnumerable<VeritradeRow> rows, string cargaId, DateTime cargaFecha)
        {
           
        // VALIDACION DUPLICIDAD: tomar “cabecera” del Excel (una sola vez)
        var first = rows.FirstOrDefault();
            if (first == null) return;

            var excelKey = first.ExcelKey;
            var tipo = first.Tipo;
            var pais = first.PaisCarga;
            var partida = first.PartidaAduanera;
            var desde = first.DesdeFiltro;
            var hasta = first.HastaFiltro;

            const string sql = @"
        INSERT INTO dbo.ImportacionesAduanas
        (
          CargaId, CargaFecha, Tipo, PaisCarga,
          PartidaAduanera, DescripcionPartidaAduanera, Aduana, DUA_DAM,
          Fecha, ETA, ManifiestoNr, CodTributario,
          Importador, Exportador, EmbarcadorExportador,
          PaisdeCompra, PuertodeEmbarque, FechadeEmbarque, Marca, PaisdeEmbarque, Producto, PaisdelExportador, RegAduana1,
          KgBruto, KgNeto, Qty1, Und1, Qty2, Und2,
          US_FOB_Tot, US_CFR_Tot, US_CIF_Tot,
          PaisOrigen, Via,
          DescripcionComercial, Descripcion1, Descripcion2, Descripcion3, Descripcion4, Descripcion5,
          DesdeFiltro, HastaFiltro, ExcelKey
        )
        VALUES
        (
          @CargaId, @CargaFecha, @Tipo, @PaisCarga,
          @Partida, @Desc, @Aduana, @DUA_DAM,
          @Fecha, @ETA, @ManifiestoNr, @CodTributario,
          @Importador, @Exportador, @EmbarcadorExportador,
          @PaisdeCompra, @PuertodeEmbarque, @FechadeEmbarque, @Marca, @PaisdeEmbarque, @Producto, @PaisdelExportador, @RegAduana1,
          @KgBruto, @KgNeto, @Qty1, @Und1, @Qty2, @Und2,
          @US_FOB_Tot, @US_CFR_Tot, @US_CIF_Tot,
          @PaisOrigen, @Via,
          @DescripcionComercial, @Descripcion1, @Descripcion2, @Descripcion3, @Descripcion4, @Descripcion5,
          @DesdeFiltro, @HastaFiltro, @ExcelKey
        );";

            await using var con = new SqlConnection(_cs);
            await con.OpenAsync();

            // 1) STARTED (si ya existe ExcelKey => skip)
            var started = await TryStartExcelAsync(con, excelKey, pais, tipo, partida, desde, hasta, null, cargaId);
            if (!started)
            {
                Console.WriteLine($"[SKIP] Ya existe ExcelKey={excelKey}. Se omite para evitar duplicados.");
                return;
            }

            try
            {
      
                foreach (var r in rows)
                {
                    using var cmd = new SqlCommand(sql, con);

                    cmd.Parameters.Add("@CargaId", SqlDbType.NVarChar, 50).Value = cargaId;
                    cmd.Parameters.Add("@CargaFecha", SqlDbType.Date).Value = cargaFecha.Date;

                    cmd.Parameters.Add("@Tipo", SqlDbType.VarChar, 50).Value = (object?)r.Tipo ?? DBNull.Value;
                    cmd.Parameters.Add("@PaisCarga", SqlDbType.VarChar, 50).Value = (object?)r.PaisCarga ?? DBNull.Value;

                    cmd.Parameters.Add("@Partida", SqlDbType.NVarChar, 20).Value = (object?)r.PartidaAduanera ?? DBNull.Value;
                    cmd.Parameters.Add("@Desc", SqlDbType.NVarChar, 500).Value = (object?)r.DescripcionPartidaAduanera ?? DBNull.Value;

                    cmd.Parameters.Add("@Aduana", SqlDbType.NVarChar, 150).Value = (object?)r.Aduana ?? DBNull.Value;
                    cmd.Parameters.Add("@DUA_DAM", SqlDbType.NVarChar, 50).Value = (object?)r.DUA_DAM ?? DBNull.Value;

                    cmd.Parameters.Add("@Fecha", SqlDbType.Date).Value = (object?)r.Fecha ?? DBNull.Value;
                    cmd.Parameters.Add("@ETA", SqlDbType.Date).Value = (object?)r.ETA ?? DBNull.Value;

                    cmd.Parameters.Add("@ManifiestoNr", SqlDbType.NVarChar, 50).Value = (object?)r.ManifiestoNr ?? DBNull.Value;
                    cmd.Parameters.Add("@CodTributario", SqlDbType.NVarChar, 20).Value = (object?)r.CodTributario ?? DBNull.Value;

                    cmd.Parameters.Add("@Importador", SqlDbType.NVarChar, 250).Value = (object?)r.Importador ?? DBNull.Value;
                    cmd.Parameters.Add("@Exportador", SqlDbType.VarChar, 500).Value = (object?)r.Exportador ?? DBNull.Value;
                    cmd.Parameters.Add("@EmbarcadorExportador", SqlDbType.VarChar, 500).Value = (object?)r.EmbarcadorExportador ?? DBNull.Value;

                    cmd.Parameters.Add("@PaisdeCompra", SqlDbType.VarChar, 200).Value = (object?)r.PaisdeCompra ?? DBNull.Value;
                    cmd.Parameters.Add("@PuertodeEmbarque", SqlDbType.VarChar, 200).Value = (object?)r.PuertodeEmbarque ?? DBNull.Value;
                    cmd.Parameters.Add("@FechadeEmbarque", SqlDbType.Date).Value = (object?)r.FechadeEmbarque ?? DBNull.Value;
                    cmd.Parameters.Add("@Marca", SqlDbType.VarChar, 200).Value = (object?)r.Marca ?? DBNull.Value;
                    cmd.Parameters.Add("@PaisdeEmbarque", SqlDbType.VarChar, 200).Value = (object?)r.PaisdeEmbarque ?? DBNull.Value;
                    cmd.Parameters.Add("@Producto", SqlDbType.VarChar, 500).Value = (object?)r.Producto ?? DBNull.Value;
                    cmd.Parameters.Add("@PaisdelExportador", SqlDbType.VarChar, 200).Value = (object?)r.PaisdelExportador ?? DBNull.Value;
                    cmd.Parameters.Add("@RegAduana1", SqlDbType.VarChar, 200).Value = (object?)r.RegAduana1 ?? DBNull.Value;

                    cmd.Parameters.Add("@KgBruto", SqlDbType.Decimal).Value = (object?)r.KgBruto ?? DBNull.Value;
                    cmd.Parameters.Add("@KgNeto", SqlDbType.Decimal).Value = (object?)r.KgNeto ?? DBNull.Value;
                    cmd.Parameters.Add("@Qty1", SqlDbType.Decimal).Value = (object?)r.Qty1 ?? DBNull.Value;
                    cmd.Parameters.Add("@Und1", SqlDbType.VarChar, 20).Value = (object?)r.Und1 ?? DBNull.Value;
                    cmd.Parameters.Add("@Qty2", SqlDbType.Decimal).Value = (object?)r.Qty2 ?? DBNull.Value;
                    cmd.Parameters.Add("@Und2", SqlDbType.VarChar, 20).Value = (object?)r.Und2 ?? DBNull.Value;

                    cmd.Parameters.Add("@US_FOB_Tot", SqlDbType.Decimal).Value = (object?)r.US_FOB_Tot ?? DBNull.Value;
                    cmd.Parameters.Add("@US_CFR_Tot", SqlDbType.Decimal).Value = (object?)r.US_CFR_Tot ?? DBNull.Value;
                    cmd.Parameters.Add("@US_CIF_Tot", SqlDbType.Decimal).Value = (object?)r.US_CIF_Tot ?? DBNull.Value;

                    cmd.Parameters.Add("@PaisOrigen", SqlDbType.VarChar, 200).Value = (object?)r.PaisOrigen ?? DBNull.Value;
                    cmd.Parameters.Add("@Via", SqlDbType.VarChar, 200).Value = (object?)r.Via ?? DBNull.Value;

                    cmd.Parameters.Add("@DescripcionComercial", SqlDbType.VarChar, 300).Value = (object?)r.DescripcionComercial ?? DBNull.Value;
                    cmd.Parameters.Add("@Descripcion1", SqlDbType.VarChar, 500).Value = (object?)r.Descripcion1 ?? DBNull.Value;
                    cmd.Parameters.Add("@Descripcion2", SqlDbType.VarChar, 500).Value = (object?)r.Descripcion2 ?? DBNull.Value;
                    cmd.Parameters.Add("@Descripcion3", SqlDbType.VarChar, 500).Value = (object?)r.Descripcion3 ?? DBNull.Value;
                    cmd.Parameters.Add("@Descripcion4", SqlDbType.VarChar, 500).Value = (object?)r.Descripcion4 ?? DBNull.Value;
                    cmd.Parameters.Add("@Descripcion5", SqlDbType.VarChar, 500).Value = (object?)r.Descripcion5 ?? DBNull.Value;

                    cmd.Parameters.Add("@DesdeFiltro", SqlDbType.VarChar, 30).Value = (object?)r.DesdeFiltro ?? DBNull.Value;
                    cmd.Parameters.Add("@HastaFiltro", SqlDbType.VarChar, 30).Value = (object?)r.HastaFiltro ?? DBNull.Value;
                    cmd.Parameters.Add("@ExcelKey", SqlDbType.VarChar, 200).Value = (object?)r.ExcelKey ?? DBNull.Value;

                    await cmd.ExecuteNonQueryAsync();
                }

                await FinishExcelAsync(con, excelKey, "OK", null);
            }
            catch (Exception ex)
            {
                await FinishExcelAsync(con, excelKey, "ERROR", ex.Message);
                throw;
            }
        }

        public async Task<bool> ExcelKeyExistsAsync(string excelKey)
        {
            const string sql = @"SELECT 1 FROM dbo.VeritradeExcelControl WHERE ExcelKey = @ExcelKey;";

            await using var con = new SqlConnection(_cs);
            await con.OpenAsync();

            await using var cmd = new SqlCommand(sql, con);
            cmd.Parameters.Add("@ExcelKey", SqlDbType.VarChar, 200).Value = excelKey;

            var obj = await cmd.ExecuteScalarAsync();
            return obj != null;
        }

        //VALIDACION DUPLICIDAD
        public async Task<bool> TryStartExcelAsync(
            SqlConnection con,
            string excelKey,
            string paisCarga,
            string tipo,
            string partida,
            string desdeFiltro,
            string hastaFiltro,
            string fileName,
            string cargaId)
        {
            const string sql = @"
              INSERT INTO dbo.VeritradeExcelControl
              (ExcelKey, PaisCarga, Tipo, Partida, DesdeFiltro, HastaFiltro, FileName, CargaId, Status)
              SELECT @ExcelKey, @PaisCarga, @Tipo, @Partida, @DesdeFiltro, @HastaFiltro, @FileName, @CargaId, 'STARTED'
              WHERE NOT EXISTS (SELECT 1 FROM dbo.VeritradeExcelControl WHERE ExcelKey = @ExcelKey);";

            using var cmd = new SqlCommand(sql, con);
            cmd.Parameters.Add("@ExcelKey", SqlDbType.VarChar, 200).Value = excelKey;
            cmd.Parameters.Add("@PaisCarga", SqlDbType.VarChar, 50).Value = paisCarga;
            cmd.Parameters.Add("@Tipo", SqlDbType.VarChar, 50).Value = tipo;
            cmd.Parameters.Add("@Partida", SqlDbType.VarChar, 20).Value = partida;
            cmd.Parameters.Add("@DesdeFiltro", SqlDbType.VarChar, 30).Value = desdeFiltro;
            cmd.Parameters.Add("@HastaFiltro", SqlDbType.VarChar, 30).Value = hastaFiltro;
            cmd.Parameters.Add("@FileName", SqlDbType.VarChar, 260).Value = (object?)fileName ?? DBNull.Value;
            cmd.Parameters.Add("@CargaId", SqlDbType.NVarChar, 50).Value = cargaId;

            var affected = await cmd.ExecuteNonQueryAsync();
            return affected == 1; // 1 = se registró (no existía); 0 = ya existía (duplicado)
        }

        public async Task FinishExcelAsync(SqlConnection con, string excelKey, string status, string? errorMsg = null)
        {
            const string sql = @"
              UPDATE dbo.VeritradeExcelControl
              SET Status = @Status,
                  ErrorMsg = @ErrorMsg
              WHERE ExcelKey = @ExcelKey;";

            using var cmd = new SqlCommand(sql, con);
            cmd.Parameters.Add("@ExcelKey", SqlDbType.VarChar, 200).Value = excelKey;
            cmd.Parameters.Add("@Status", SqlDbType.VarChar, 20).Value = status;
            cmd.Parameters.Add("@ErrorMsg", SqlDbType.VarChar, 4000).Value = (object?)errorMsg ?? DBNull.Value;

            await cmd.ExecuteNonQueryAsync();
        }

    }
}


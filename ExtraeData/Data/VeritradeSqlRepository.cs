using System;
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

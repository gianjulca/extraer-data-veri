using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtraeData.Models
{
    public sealed class VeritradeRow
    {
        public string PaisCarga { get; init; } = "";
        public string Tipo { get; init; } = "";
        public string PartidaAduanera { get; init; } = "";
        public string DescripcionPartidaAduanera { get; init; } = "";
        public string Aduana { get; init; } = "";
        public string DUA_DAM { get; init; } = "";
        public DateTime? Fecha { get; init; }
        public DateTime? ETA { get; init; }
        public string ManifiestoNr { get; init; } = "";
        public string CodTributario { get; init; } = "";
        public string Importador { get; init; } = "";
        public string Exportador { get; init; } = "";
        public string PaisdeCompra { get; init; } = "";
        public string PuertodeEmbarque { get; init; } = "";
        public DateTime? FechadeEmbarque { get; init; }
        public string Marca { get; init; } = "";
        public string PaisdeEmbarque { get; init; } = "";
        public string Producto { get; init; } = "";
        public string PaisdelExportador { get; init; } = "";
        public string RegAduana1 { get; init; } = "";
        public string EmbarcadorExportador { get; init; } = "";
        public decimal? KgBruto { get; init; }
        public decimal? KgNeto { get; init; }
        public decimal? Qty1 { get; init; }
        public string Und1 { get; init; } = "";
        public decimal? Qty2 { get; init; }
        public string Und2 { get; init; } = "";
        public decimal? US_FOB_Tot { get; init; }
        public decimal? US_CFR_Tot { get; init; }
        public decimal? US_CIF_Tot { get; init; }
        public string PaisOrigen { get; init; } = "";
        public string Via { get; init; } = "";
        public string DescripcionComercial { get; init; } = "";
        public string Descripcion1 { get; init; } = "";
        public string Descripcion2 { get; init; } = "";
        public string Descripcion3 { get; init; } = "";
        public string Descripcion4 { get; init; } = "";
        public string Descripcion5 { get; init; } = "";

        //VALIDACION DE DUPLICIDAD
        public string DesdeFiltro { get; init; } = "";
        public string HastaFiltro { get; init; } = "";
        public string ExcelKey { get; init; } = "";
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class CortesDetalle
    {
        public int iCodCorteDetalle { get; set; }

        public int iCodCorte { get; set; }

        public int idias { get; set; }

        public string vMensaje { get; set; }

        public int piPageSize { get; set; }

        public int piCurrentPage { get; set; }

        public string pvSortColumn { get; set; }

        public string pvSortOrder { get; set; }

        public int totalRegistros { get; set; }

        public int totalPaginas { get; set; }

        public int paginaActual { get; set; }

        public string Entregable { get; set; }

    }
}
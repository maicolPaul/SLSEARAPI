using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class Cronograma
    {
        public int iCodExtensionista { get; set; }

        public int iCodCronograma { get; set; }

        public int iCodComponente { get; set; }

        public int iCodActividad { get; set; }

        public int iCantidad { get; set; }

        public string dFecha { get; set; }

        public int iopcion { get; set; }

        public string vMensaje { get; set; }

        public int piPageSize { get; set; }

        public int piCurrentPage { get; set; }


        public string pvSortColumn { get; set; }

        public string pvSortOrder { get; set; }

        public int iCodIdentificacion { get; set; }

        public string vDescripcionActividad { get; set; }

        public string vActividad { get; set; }

        public string vUnidadMedida { get; set; }

        public string vMeta { get; set; }

        public int totalRegistros { get; set; }
        public int totalPaginas { get; set; }

        public int paginaActual { get; set; }

        public int nTipoActividad { get; set; }


    }
}
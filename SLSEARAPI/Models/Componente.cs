using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


namespace SLSEARAPI.Models
{
    public class Componente
    {
        public int iCodComponente { get; set; }

        public string vDescripcion { get; set; }
        public string vIndicador { get; set; }

        public string vUnidadMedida { get; set; }

        public string vMeta { get; set; } 

        public string vMedio { get; set; }

        public int nTipoComponente { get; set; }

        public int iCodIdentificacion { get; set; }

        public string vDescComponente { get; set; }


        public int piPageSize { get; set; }

        public int piCurrentPage { get; set; }

        public string pvSortColumn { get; set; }
        public string pvSortOrder { get; set; }
        public int totalRegistros { get; set; }

        public string vMensaje { get; set; }
    }
}
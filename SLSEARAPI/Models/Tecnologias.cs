using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class Tecnologias
    {
        public int iCodTecnologia { get; set; }
        public int iCodIdentificacion { get; set; }
        public string vtecnologia1 { get; set; }
        public string vtecnologia2 { get; set; }
        public string vtecnologia3 { get; set; }

        public string vMensaje { get; set; }

        public int piPageSize { get; set; }

        public int piCurrentPage { get; set; }

        public string pvSortColumn { get; set; }

        public string pvSortOrder { get; set; }

        public int totalRegistros { get; set; }

    }
}
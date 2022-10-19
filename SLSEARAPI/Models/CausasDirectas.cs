using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class CausasDirectas
    {
        public int iCodCausaDirecta { get; set; }

        public string vdescrcausadirecta { get; set; }

        public int iCodIdentificacion { get; set; }

        public string vMensaje { get; set; }

        public int piPageSize { get; set; }

        public int piCurrentPage { get; set; }

        public string pvSortColumn { get; set; }

        public string pvSortOrder { get; set; }

        public int totalRegistros { get; set; }



    }
}
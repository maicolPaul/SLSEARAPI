using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class PorductorEjecucionTecnica
    {
        public int iCodProEje { get; set; }

        public int iCodComponente { get; set; }

        public int iCodActividad { get; set; }

        public int iCodProductor { get; set; }

        public string dFechaCapa { get; set; }

        public string vTipo { get; set; }

        public string vMensaje { get; set; }
    }
}
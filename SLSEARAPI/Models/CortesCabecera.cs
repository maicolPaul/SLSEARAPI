using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class CortesCabecera
    {
        public int iCodCorte { get; set; }

        public int iCodFichaTecnica { get; set; }

        public string dFechaInicioReal { get; set; }

        public string dFechaFinReal { get; set; }

        public string vMensaje { get; set; }
    }
}
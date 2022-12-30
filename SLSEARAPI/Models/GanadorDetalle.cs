using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class GanadorDetalle
    {
        public int iCodganadordetalle { get; set; }

        public int iCodganador { get; set; }

        public int iCodEvaluador { get; set; }

        public string vMensaje { get; set; }
    }
}
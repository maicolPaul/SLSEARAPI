using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class Indicadores
    {
        public int iCodIndicador { get; set; }
        public int iCodIdentificacion { get; set; }

        public int iCodigoIdentificador { get; set; }

        public string vDescIdentificador { get; set; }

        public string vUnidadMedida { get; set; }

        public string vMeta { get; set; }

        public string vMedioVerificacion { get; set; }


    }
}
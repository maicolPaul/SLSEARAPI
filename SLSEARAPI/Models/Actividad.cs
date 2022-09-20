using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class Actividad
    {
        public int iCodActividad { get; set; }

        public int iCodIdentificacion { get; set; }

        public string vActividad { get; set; }

        public string vDescripcion { get; set; }

        public string vUnidadMedida { get; set; }

        public string vMeta { get; set; }

        public string vMedio { get; set; }

        public int nTipoActividad { get; set; }

        public int iCodExtensionista { get; set; }

        public int resumen { get; set; }
    }
}


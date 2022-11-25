using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class SupervisionCapCab
    {
        public int iCodSuperCab { get; set; }

        public int iCodIdentificacion { get; set; }

        public int iCodFichaTecnica { get; set; }

        public int iCodComponente { get; set; }

        public int iCodActividad { get; set; }

        public string vObservaciongeneral { get; set; }

        public string vRecomendacion { get; set; }

        public string vNombreSupervisor { get; set; }

        public string vCargoSupervisor { get; set; }

        public string vEntidadSupervisor { get; set; }

        public int iCodCalificacion { get; set; }

        public string vMensaje { get; set; }

    }
}
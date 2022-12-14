using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class PlanCapacitacion
    {
        public int iopcion { get; set; }
        public string vMensaje { get; set; }
        public int piPageSize { get; set; }
        public int piCurrentPage { get; set; }
        public string pvSortColumn { get; set; }
        public string pvSortOrder { get; set; }
        public int iCodPlanCap { get; set; }
        public int iCodActividad { get; set; }
        public string vModuloTema { get; set; }
        public string vObjetivo { get; set; }
        public int iMeta { get; set; }
        public int iBeneficiario { get; set; }
        public string dFechaActividad { get; set; }
        public decimal iTotalTeoria { get; set; }
        public decimal iTotalPractica { get; set; }
        public bool bActivo { get; set; }
        public int iRecordCount { get; set; }
        public int iCodHito { get; set; }

        public int iCodExtensionista { get; set; }

        public string porcentaje { get; set; }
    }
}
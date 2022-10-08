using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Web;

namespace SLSEARAPI.Models
{
    public class PlanSesion
    {
        public int iopcion { get; set; }
        public string vMensaje { get; set; }
        public int piPageSize { get; set; }
        public int piCurrentPage { get; set; }
        public string pvSortColumn { get; set; }
        public string pvSortOrder { get; set; }
        public int iCodPlanSesion { get; set; }
        public int iCodPlanCap { get; set; }
        public int iDuracion { get; set; }
        public string vTematica { get; set; }
        public string vDescripMetodologia { get; set; }
        public string vMateriales { get; set; }
        public bool bActivo { get; set; }
        public int iRecordCount { get; set; }
    }
}
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class Costo
    {
        public int iopcion { get; set; }
        public string vMensaje { get; set; }
        public int piPageSize { get; set; }
        public int piCurrentPage { get; set; }
        public string pvSortColumn { get; set; }
        public string pvSortOrder { get; set; }

        public int iCodCosto { get; set; }
        public int iCodIdentificacion { get; set; }
        public int iCodExtensionista { get; set; }
        public int iCodComponente { get; set; }
        public string vComponente { get; set; }
        public int iCodActividad { get; set; }
        public string vActividad { get; set; }
        public int iTipoMatServ { get; set; }
        public string TipoMatServ { get; set; }
        public string vDescripcion { get; set; }
        public string vUnidadMedida { get; set; }
        public int iCantidad { get; set; }
        public decimal dCostoUnitario { get; set; }
        public string dFecha { get; set; }
        public string dFechaRegistro { get; set; }
        public bool bActivo { get; set; }
        public string Estado { get; set; }
        public int iRecordCount { get; set; }

        public int iCodHito { get; set; }

    }
}
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
        public string vDescripcion_ { get; set; }
        public string vUnidadMedida { get; set; }
        public string vMeta { get; set; }
        public string vMedio { get; set; }
        public string vMedio_ { get; set; }
        public int nTipoActividad { get; set; }
        public string vDescripcionCorta { get; set; }
        public int Correlativo { get; set; }
        public int iCodExtensionista { get; set; }
        public int resumen { get; set; }
        public int iopcion { get; set; }
        public string vMensaje { get; set; }
        public int piPageSize { get; set; }
        public int piCurrentPage { get; set; }
        public string pvSortColumn { get; set; }
        public string pvSortOrder { get; set; }
        public bool bActivo { get; set; }
        public int iRecordCount { get; set; }
        public int totalRegistros { get; set; }
        public string vActividadCorrelativo { get; set; }

        public int iCodComponenteDesc { get; set; }

        public string vUnidadMedidaCorta { get; set; }
        public string vMetaCorta { get; set; }
        public string vMedioCorta { get; set; }
        public string dFecha { get; set; }
        public string dFechaFin { get; set; }
    }
}


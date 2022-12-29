using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class FichaEvaluacion
    {
        public int piPageSize { get; set; }
        public int piCurrentPage { get; set; }
        public string pvSortColumn { get; set; }
        public string pvSortOrder { get; set; }
        public int iCodFichaEvaluacion { get; set; }
        public int iCodIdentificacion { get; set; }
        public int iCodComiteEvaluador { get; set; }
        public int iCodCategoria { get; set; }
        public string vCategoria { get; set; }
        public int iCodCriterio { get; set; }
        public string vCriterio { get; set; }
        public string vCriterio1 { get; set; }
        public decimal PuntajeMaximo { get; set; }
        public decimal dPuntajeEvaluacion { get; set; }
        public string vJustificacion { get; set; }
        public int iRecordCount { get; set; }
        public int iOpcion { get; set; }
        public string vMensaje { get; set; }

        public int iCodComiteIdentificacion { get; set; }
    }
}
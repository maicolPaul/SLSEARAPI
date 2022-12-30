using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class ComiteIdentificacion
    {
        public int piPageSize { get; set; }
        public int piCurrentPage { get; set; }
        public string pvSortColumn { get; set; }
        public string pvSortOrder { get; set; }
        public int iCodComiteIdentificacion { get; set; }
        public int iCodComiteEvaluador { get; set; }
        public int iCodIdentificacion { get; set; }
        public string dFechaRegistro { get; set; }
        public bool bActivo { get; set; }
        public string vNombreSearT1 { get; set; }
        public string vDireccionT2 { get; set; }
        public string iCodUbigeoT1 { get; set; }
        public string vNomDepartamento { get; set; }
        public string vNomProvincia { get; set; }
        public string vNomDistrito { get; set; }
        public int iRecordCount { get; set; }
        public string vMensaje { get; set; }
        public int iCodExtensionista { get; set; }
        public int EvaluadoFinalizar { get; set; }
        public string vCodDepartamento { get; set; }
        public decimal dPuntajeEvaluacion { get; set; }
    }
}
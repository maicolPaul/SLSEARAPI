using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class ComiteEvaluador
    {
        public int iCodComiteEvaluador { get; set; }

        public int iCodIdentificacion { get; set; }

        public string vNombres { get; set; }

        public string vApellidoPat { get; set; }

        public string vApellidoMat { get; set; }

        public int iCodTipoDoc { get; set; }

        public string vNroDocumento { get; set; }

        public string vCodUbigeo { get; set; }

        public string vNomDistrito { get; set; }

        public string vCodProvincia { get; set; }

        public string vNomProvincia { get; set; }

        public string vCodDepartamento { get; set; }

        public string vNomDepartamento { get; set; }

        public int iCodCargo { get; set; }

        public string vDescripcionCargo { get; set; }

        public string vCelular { get; set; }

        public string vCorreo { get; set; }

        public int iCodArchivos { get; set; }

        public int iopcion { get; set; }

        public string vMensaje { get; set; }

        public int piPageSize { get; set; }

        public int piCurrentPage { get; set; }

        public int pvSortColumn { get; set; }

        public int pvSortOrder { get; set; }

        public int totalRegistros { get; set; }

        public string Estado { get; set; }
        public string vContrasena { get; set; }
    }
}
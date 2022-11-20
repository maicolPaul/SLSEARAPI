using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class Productor
    {
        public int iCodProductor { get; set; }
        public Int64 Nro { get; set; }
        public string vApellidosNombres { get; set; }
        public string vDni { get; set; }
        public string vCelular { get; set; }
        public int iEdad { get; set; }
        public int iSexo { get; set; }
        public string vSexo { get; set; }
        public string vRecibioCapacitacion { get; set; }
        public int iPerteneceOrganizacion { get; set; }
        public string vNombreOrganizacion { get; set; }
        public int iRecibioCapacitacion { get; set; }
        public int iEsRepresentante { get; set; }
        public int iCodExtensionista { get; set; }
        public string vMensaje { get; set; }
        public int piPageSize { get; set; }
        public int piCurrentPage { get; set; }
        public string pvSortColumn { get; set; }
        public string pvSortOrder { get; set; }
        public int totalRegistros { get; set; }
        public int totalPaginas { get; set; }

        public int paginaActual { get; set; }

        public int iOpcion { get; set; }

        public string vNombreRepresentante { get; set; }

        public string vRucOrg { get; set; }

        public string vTelefonoOrg { get; set; }

        public string vCelularOrg { get; set; }

        public string vDireccionOrg { get; set; }

        public string vCorreoElectronicoOrg { get; set; }

        public int iCodTipoOrg { get; set; }
        public string vOrganizacion { get; set; }

        public int cantidadmasculino { get; set; }

        public int cantidadfemenino { get; set; }

        public decimal promedio { get; set; }

        public int jovenes { get; set; }

        public int recibiocapacitacion { get; set; }

        public decimal porfemenino { get; set; }

        public decimal porjovenes { get; set; }

        public decimal porrecibiocapacitacion { get; set; }


    }
}


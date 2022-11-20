using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class FichaTecnica
    {
        /// <summary>
        ///  1.1
        /// </summary>
        public int iCodFichaTecnica { get; set; }

        public string vNombreSearT1 { get; set; }

        public string vNaturalezaIntervencionT1 { get; set; }

        public string vSubSectorT1 { get; set; }

        public string vCadenaProductivaT1 { get; set; }

        public string vProcesoProductivaT1 { get; set; }

        public string vLineaPrioritariaT1 { get; set; }

        public string vProductoServicioAmpliarT1 { get; set; }

        public string iCodUbigeoT1 { get; set; }

        public string vNomDepartamento { get; set; }
        public string vNomProvincia { get; set; }
        public string vNomDistrito { get; set; }

        public string vLocalidadT1 { get; set; }

        public string vZonaUTMT1 { get; set; }

        public string vCoordenadasUTMNorteT1 { get; set; }

        public string vCoordenadasUTMEsteT1 { get; set; }

        public string dFechaInicioServicioT1 { get; set; }

        public string dFechaFinServicioT1 { get; set; }
        public int TotalDias { get; set; }

        public int iDuracionT1 { get; set; }

        /// <summary>
        /// 1.2
        /// </summary>
        

        public string vNombreEntidadProponenteT2 { get; set; }

        public string vNombreDireccionPerteneceT2 { get; set; }

        public string vDireccionT2 { get; set; }

        public string vTelefonoT2 { get; set; }
        public string TipoPersoneriaT3 { get; set; }

        public string vCorreoElectronicoT2 { get; set; }

        public string vNombreDirectorAgenciaAgrariaT2 { get; set; }

        public string vDireccionZonaAgroruralT2 { get; set; }

        /// <summary>
        /// 1.3
        /// </summary>

        public int iCodTipoPersoneriaT3 { get; set; }

        public string vNombreRazonSocialProveedorT3 { get; set; }

        public string vNombreRepresentanteLegalT3 { get; set; }

        public string vDniT3 { get; set; }

        public string vRucT3 { get; set; }

        public string vDireccionT3 { get; set; }

        public string vTelefonoT3 { get; set; }
        public string vCelularT3 { get; set; }
        public string vCorreoElectronicoT3 { get; set; }
        public string vPaginaWebT3 { get; set; }
        public string vEpecialidadProveedorT3 { get; set; }
        public int iCodTipoProveedorT3 { get; set; }
        public string vProveedor { get; set; }
        public int iCodConvocatoria { get; set; }

        public int iCodExtensionista { get; set; }

        public string vMensaje { get; set; }

        public int error  { get; set; }

        public int iOpcion { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class Identificacion
    {
        public int iCodIdentificacion { get; set; }

        public string vLimitaciones { get; set; }

        public string vEstadoSituacional { get; set; }

        public int iCodExtensionista { get; set; }

        public int iOpcion { get; set; }

        public string vMensaje { get; set; }

        public List<Tecnologias> listatecnologias { get; set; }

        public string tecnologiasxml { get; set; }
        public List<Indicadores> listaindicadores { get; set; }
        public List<CausasDirectas> listacausasdirectas { get; set; }        
        public List<CausasIndirectas> listacausasindirectas { get; set; }
        public List<EfectoDirecto> listaefectodirectos { get; set; }

        public List<EfectoIndirecto> listaefectoindirectos { get; set; }

        public List<Componente> listacomponente { get; set; }
        public List<Actividad> listaactividad { get; set; }
        
        public string causasdirectasxml { get; set; }
        public string causasindirectaxml { get; set; }
        public string indicadoresxml { get; set; }
        public string efectosdirectosxml { get; set; }
        public string efectosindirectosxml { get; set; }
        public string componentesxml { get; set; }

        public string actividadesxml { get; set; }
        public string vProblemacentral { get; set; }
        public string vNumeroUnidadesProductivas { get; set; }
        public string vUnidadMedidaProductivas { get; set; }
        public int vNumerosFamiliares { get; set; }
        public  int vCantidad { get; set; }

        public string vUnidadMedida { get; set; }
        public string vRendimientoCadenaProductiva { get; set; }
        public string vGremios { get; set; }
        public string vObjetivoCentral { get; set; }

        public string vDescComponente1 { get; set; }
        public string vDescComponente2 { get; set; }

        public int piPageSize { get; set; }

        public int piCurrentPage { get; set; }

        public string pvSortColumn { get; set; }

        public string pvSortOrder { get; set; }

        public int totalRegistros { get; set; }
        public string vNombreSearT1 { get; set; }
        public string iCodUbigeoT1 { get; set; }
        public string vEstado { get; set; }

        public int cantidadproductores { get; set; }
    }
}
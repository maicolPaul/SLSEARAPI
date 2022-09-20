using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace SLSEARAPI.Controllers
{
    public class TipoEntidadController : ApiController
    {
        TipoEntidadDL tipoEntidadDL;

        private Exception ex;
        public TipoEntidadController()
        {

            tipoEntidadDL = new TipoEntidadDL();
        }

        [HttpPost]
        [ActionName("ListarTipoEntidad")]

        public List<TipoEntidad> ListarTipoEntidad(TipoEntidad TipoEntidad)
        {
            try
            {
                return tipoEntidadDL.ListarTipoEntidad(TipoEntidad);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

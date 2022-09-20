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
    public class NivelDeInstruccionController : ApiController
    {
        NivelDeInstruccionDL nivelDeInstruccionDL;

        private Exception ex;
        public NivelDeInstruccionController()
        {

            nivelDeInstruccionDL = new NivelDeInstruccionDL();
        }

        [HttpPost]
        [ActionName("ListarNivelDeInstruccion")]

        public List<NivelDeInstruccion> ListarNivelDeInstruccion(NivelDeInstruccion NivelDeInstruccion)
        {
            try
            {
                return nivelDeInstruccionDL.ListarNivelDeInstruccion(NivelDeInstruccion);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

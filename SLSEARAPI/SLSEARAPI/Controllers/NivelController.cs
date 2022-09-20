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
    public class NivelController : ApiController
    {

        NivelDL nivelDL;

        private Exception ex;
        public NivelController()
        {

            nivelDL = new NivelDL();
        }

        [HttpPost]
        [ActionName("ListarNivel")]

        public List<Nivel> ListarNivel(Nivel Nivel)
        {
            try
            {
                return nivelDL.ListarNivel(Nivel);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

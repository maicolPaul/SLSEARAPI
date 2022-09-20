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
    public class RequisitosController : ApiController
    {
        RequisitosDL requisitosDL;

        private Exception ex;
        public RequisitosController()
        {

            requisitosDL = new RequisitosDL();
        }

        [HttpPost]
        [ActionName("ListarRequisitos")]

        public List<Requisitos> ListarRequisitos(Requisitos Requisitos)
        {
            try
            {
                return requisitosDL.ListarRequisitos(Requisitos);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

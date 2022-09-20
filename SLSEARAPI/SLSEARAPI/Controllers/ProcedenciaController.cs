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
    public class ProcedenciaController : ApiController
    {
        ProcedenciaDL procedenciaDL;

        private Exception ex;
        public ProcedenciaController()
        {

            procedenciaDL = new ProcedenciaDL();
        }

        [HttpPost]
        [ActionName("ListarProcedencia")]

        public List<Procedencia> ListarProcedencia(Procedencia Procedencia)
        {
            try
            {
                return procedenciaDL.ListarProcedencia(Procedencia);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

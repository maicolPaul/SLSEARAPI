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
    public class AgenciaAgrariaController : ApiController
    {
        AgenciaAgrariaDL agenciaAgrariaDL;

        private Exception ex;
        public AgenciaAgrariaController()
        {

            agenciaAgrariaDL = new AgenciaAgrariaDL();
        }

        [HttpPost]
        [ActionName("ListarAgenciaAgraria")]

        public List<AgenciaAgraria> ListarAgenciaAgraria(AgenciaAgraria AgenciaAgraria)
        {
            try
            {
                return agenciaAgrariaDL.ListarAgenciaAgraria(AgenciaAgraria);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

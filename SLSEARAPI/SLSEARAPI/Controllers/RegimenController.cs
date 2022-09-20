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
    public class RegimenController : ApiController
    {
        RegimenDL regimenDL;

        private Exception ex;
        public RegimenController()
        {

            regimenDL = new RegimenDL();
        }

        [HttpPost]
        [ActionName("ListarRegimen")]

        public List<Regimen> ListarRegimen(Regimen Regimen)
        {
            try
            {
                return regimenDL.ListarRegimen(Regimen);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

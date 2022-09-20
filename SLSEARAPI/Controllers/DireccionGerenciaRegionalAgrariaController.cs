using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Web.Http;

namespace SLSEARAPI.Controllers
{
    [DataContract]
    public class DireccionGerenciaRegionalAgrariaController : ApiController
    {
        DireccionGerenciaRegionalAgrariaDL direccionGerenciaRegionalAgrariaDL;

        private Exception ex;
        public DireccionGerenciaRegionalAgrariaController()
        {

            direccionGerenciaRegionalAgrariaDL = new DireccionGerenciaRegionalAgrariaDL();
        }

        [HttpPost]
        [ActionName("ListarDireccionGerenciaRegionalAgraria")]

        public List<DireccionGerenciaRegionalAgraria> ListarDireccionGerenciaRegionalAgraria(DireccionGerenciaRegionalAgraria DireccionGerenciaRegionalAgraria)
        {
            try
            {
                return direccionGerenciaRegionalAgrariaDL.ListarDireccionGerenciaRegionalAgraria(DireccionGerenciaRegionalAgraria);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

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
    public class AsignacionEvaluadorController : ApiController
    {
        AsignacionEvaluadorDL asignacionEvaluadorDL;
        public AsignacionEvaluadorController()
        {
            asignacionEvaluadorDL = new AsignacionEvaluadorDL();
        }

        [HttpPost]
        [ActionName("ListarSear")]
        public List<Identificacion> ListarSear(Identificacion identificacion)
        {
            try
            {
                return asignacionEvaluadorDL.ListarSear(identificacion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

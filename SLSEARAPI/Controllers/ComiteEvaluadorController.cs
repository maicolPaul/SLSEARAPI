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
    public class ComiteEvaluadorController : ApiController
    {
        ComiteEvaluadorDL comiteEvaluadorDL;
        public ComiteEvaluadorController()
        {
            comiteEvaluadorDL = new ComiteEvaluadorDL();
        }

        [HttpPost]
        [ActionName("InsertarComiteEvaluador")]
        public ComiteEvaluador InsertarComiteEvaluador(ComiteEvaluador comiteEvaluador)
        {
            try
            {
                return comiteEvaluadorDL.InsertarComiteEvaluador(comiteEvaluador);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarComiteEvaluador")]
        public List<ComiteEvaluador> ListarComiteEvaluador(ComiteEvaluador comiteEvaluador)
        {
            try
            {
                return comiteEvaluadorDL.ListarComiteEvaluador(comiteEvaluador);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        [HttpPost]
        [ActionName("ListarCargo")]
        public List<Cargo> ListarCargo()
        {
            try
            {
                return comiteEvaluadorDL.ListarCargo();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

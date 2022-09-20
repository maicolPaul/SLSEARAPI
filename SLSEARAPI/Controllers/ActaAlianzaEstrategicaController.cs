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
    public class ActaAlianzaEstrategicaController : ApiController
    {
        ActaAlianzaEstrategicaDL actaAlianzaEstrategicaDL;
        public ActaAlianzaEstrategicaController()
        {
            actaAlianzaEstrategicaDL = new ActaAlianzaEstrategicaDL();
        }
        //Metodo Insertar Productor
        [HttpPost]
        [ActionName("InsertarProductor")]
        public Productor InsertarProductor(Productor entidad)
        {
            try
            {
                return actaAlianzaEstrategicaDL.InsertarProductor(entidad);
            }
            catch (Exception)
            {
                throw;
            }
        }
        //Metodo Listar Productor
        [HttpPost]
        [ActionName("ListarProductor")]
        public List<Productor> ListarProductor(Productor entidad)
        {
            try
            {
                return actaAlianzaEstrategicaDL.ListarProductor(entidad);
            }
            catch (Exception)
            {
                throw;
            }
        }
        [HttpPost]
        [ActionName("ListarRepresentantes")]
        public List<Productor> ListarRepresentantes(Productor entidad)
        {
            try
            {
                return actaAlianzaEstrategicaDL.ListarRepresentantes(entidad);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

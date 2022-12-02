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
    public class HitoController : ApiController
    {
        HitosDL hitosDL;

        private Exception ex;
        public HitoController()
        {

            hitosDL = new HitosDL();
        }

        [HttpPost]
        [ActionName("InsertarHito")]
        public Hito InsertarHito(Hito entidad)
        {
            try
            {
                return hitosDL.InsertarHito(entidad);
            }
            catch (Exception)
            {
                throw;
            }

        }
        [HttpPost]
        [ActionName("InsertarProductorEje")]
        public PorductorEjecucionTecnica InsertarProductorEje(PorductorEjecucionTecnica porductorEjecucionTecnica)
        {
            try
            {
                return hitosDL.InsertarProductorEje(porductorEjecucionTecnica);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
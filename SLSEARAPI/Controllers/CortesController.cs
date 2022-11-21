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
    public class CortesController : ApiController
    {
        CortesDL cortesDL;
        public CortesController()
        {
            cortesDL=new CortesDL();
        }

        [HttpPost]
        [ActionName("InsertarCorteCabecera")]
        public CortesCabecera InsertarCorteCabecera(CortesCabecera cortesCabecera)
        {
            try
            {
                return cortesDL.InsertarCorteCabecera(cortesCabecera);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarCorteDetalle")]
        public CortesDetalle InsertarCorteDetalle(CortesDetalle cortesDetalle)
        {
            try
            {
                return cortesDL.InsertarCorteDetalle(cortesDetalle);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarCortesDetalle")]
        public List<CortesDetalle> ListarCortesDetalle(CortesDetalle cortesDetalle)
        {
            try
            {
                return cortesDL.ListarCortesDetalle(cortesDetalle);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EliminarCorteDetalle")]
        public CortesDetalle EliminarCorteDetalle(CortesDetalle cortesDetalle)
        {
            try
            {
                return cortesDL.EliminarCorteDetalle(cortesDetalle);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ObtenerCorteCabecera")]
        public CortesCabecera ObtenerCorteCabecera(CortesCabecera cortesCabecera)
        {
            try
            {
                return cortesDL.ObtenerCorteCabecera(cortesCabecera);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

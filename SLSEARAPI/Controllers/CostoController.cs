using OfficeOpenXml;
using OfficeOpenXml.Style;
using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace SLSEARAPI.Controllers
{
    public class CostoController : ApiController
    {
        CostoDL costoDL;
        public CostoController()
        {
            costoDL = new CostoDL();
        }

        [HttpPost]
        [ActionName("ListarCosto")]
        public List<Costo> ListarCosto(Costo costo)
        {
            try
            {
                return costoDL.ListarCosto(costo);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarActividad")]
        public List<Actividad> ListarActividad(Actividad actividad)
        {
            try
            {
                return costoDL.ListarActividad(actividad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarCosto")]
        public Costo InsertarCosto(Costo costo)
        {
            try
            {
                return costoDL.InsertarCosto(costo);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
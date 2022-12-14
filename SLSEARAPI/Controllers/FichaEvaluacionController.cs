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
    public class FichaEvaluacionController : ApiController
    {
        FichaEvaluacionDL fichaEvaluacionDL;
        public FichaEvaluacionController()
        {
            fichaEvaluacionDL = new FichaEvaluacionDL();
        }

        [HttpPost]
        [ActionName("ListarComiteIdentificacion")]
        public List<ComiteIdentificacion> ListarComiteIdentificacion(ComiteIdentificacion comiteIdentificacion)
        {
            try
            {
                return fichaEvaluacionDL.ListarComiteIdentificacion(comiteIdentificacion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarFichaEvaluacion")]
        public List<FichaEvaluacion> ListarFichaEvaluacion(FichaEvaluacion fichaEvaluacion)
        {
            try
            {
                return fichaEvaluacionDL.ListarFichaEvaluacion(fichaEvaluacion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarFichaEvaluacion")]
        public FichaEvaluacion InsertarFichaEvaluacion(FichaEvaluacion fichaEvaluacion)
        {
            try
            {
                return fichaEvaluacionDL.InsertarFichaEvaluacion(fichaEvaluacion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("FinalizarEvaluacion")]
        public FichaEvaluacion FinalizarEvaluacion(FichaEvaluacion fichaEvaluacion)
        {
            try
            {
                return fichaEvaluacionDL.FinalizarEvaluacion(fichaEvaluacion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarEvaluadoFinalizados")]
        public List<ComiteIdentificacion> ListarEvaluadoFinalizados(ComiteIdentificacion comiteIdentificacion)
        {
            try
            {
                return fichaEvaluacionDL.ListarEvaluadoFinalizados(comiteIdentificacion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarGanador")]
        public Ganador InsertarGanador(Ganador ganador)
        {
            try
            {
                return fichaEvaluacionDL.InsertarGanador(ganador);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarGanadorDetalle")]
        public GanadorDetalle InsertarGanadorDetalle(GanadorDetalle ganadordetalle)
        {
            try
            {
                return fichaEvaluacionDL.InsertarGanadorDetalle(ganadordetalle);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

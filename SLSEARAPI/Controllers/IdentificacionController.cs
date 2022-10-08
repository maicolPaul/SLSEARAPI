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
    public class IdentificacionController : ApiController
    {
        IdentificacionDL identificacionDL;

        private Exception ex;
        public IdentificacionController()
        {

            identificacionDL = new IdentificacionDL();
        }

        [HttpPost]
        [ActionName("ListarIndicadores")]
        public List<Indicadores> ListarIndicadores(Indicadores indicadores)
        {
            try
            {
                return identificacionDL.ListarIndicadores(indicadores);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        [HttpPost]
        [ActionName("ListarActividades")]
        public List<Actividad> ListarActividades(Actividad actividad)
        {
            try
            {
                return identificacionDL.ListarActividades(actividad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarComponentePorUsuario")]
        public List<Componente> ListarComponentePorUsuario(Componente componente)
        {
            try
            {
                return identificacionDL.ListarComponentePorUsuario(componente);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarComponentesPaginadoPorUsuario")]
        public List<Componente> ListarComponentesPaginadoPorUsuario(Componente componente)
        {
            try
            {
                return identificacionDL.ListarComponentesPaginadoPorUsuario(componente);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        [HttpPost]
        [ActionName("InsertarIdentificacion")]
        public Identificacion InsertarIdentificacion(Identificacion identificacion)
        {
            try
            {               

                return identificacionDL.InsertarIdentificacion(identificacion);
            }
            catch (Exception)
            {
                throw;
            }
        }

        [HttpPost]
        [ActionName("ListarIdentificacion")]
        public List<Identificacion> ListarIdentificacion(Identificacion identificacion)
        {
            try
            {
                return identificacionDL.ListarIdentificacion(identificacion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        [HttpPost]
        [ActionName("ListarTecnologias")]
        public List<Tecnologias> ListarTecnologias(Tecnologias tecnologias)
        {
            try
            {
                return identificacionDL.ListarTecnologias(tecnologias);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarActividadesPorComponente")]
        public List<Actividad> ListarActividadesPorComponente(Actividad actividad)
        {
            try
            {
                return identificacionDL.ListarActividadesPorComponente(actividad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarActividad")]
        public Actividad InsertarActividad(Actividad actividad)
        {
            try
            {
                return identificacionDL.InsertarActividad(actividad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarComponente")]
        public Componente InsertarComponente(Componente componente)
        {
            try
            {
                return identificacionDL.InsertarComponente(componente);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EliminarActividad")]
        public Actividad EliminarActividad(Actividad actividad)
        {
            try
            {
                return identificacionDL.EliminarActividad(actividad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

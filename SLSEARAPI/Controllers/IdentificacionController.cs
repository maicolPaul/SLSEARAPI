using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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

        [HttpPost]
        [ActionName("ActualizarComponente")]
        public Componente ActualizarComponente(Componente componente)
        {
            try
            {
                return identificacionDL.ActualizarComponente(componente);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EliminarComponente")]
        public Componente EliminarComponente(Componente componente)
        {
            try
            {
                return identificacionDL.EliminarComponente(componente);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ActualizarActividad")]
        public Actividad ActualizarActividad(Actividad actividad)
        {
            try
            {
                return identificacionDL.ActualizarActividad(actividad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarTecnologia")]
        public Tecnologias InsertarTecnologia(Tecnologias tecnologias)
        {
            try
            {
                return identificacionDL.InsertarTecnologia(tecnologias);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EditarTecnologia")]
        public Tecnologias EditarTecnologia(Tecnologias tecnologias)
        {
            try
            {
                return identificacionDL.EditarTecnologia(tecnologias);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EliminarTecnologia")]
        public Tecnologias EliminarTecnologia(Tecnologias tecnologias)
        {
            try
            {
                return identificacionDL.EliminarTecnologia(tecnologias);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarTecnologiasPaginado")]
        public List<Tecnologias> ListarTecnologiasPaginado(Tecnologias tecnologias)
        {
            try
            {
                return identificacionDL.ListarTecnologiasPaginado(tecnologias);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarCausasDirectas")]
        public CausasDirectas InsertarCausasDirectas(CausasDirectas causasDirectas)
        {
            try
            {
                return identificacionDL.InsertarCausasDirectas(causasDirectas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ActividadCorrelativo")]
        public Actividad ActividadCorrelativo(Actividad actividad)
        {
            try
            {
                return identificacionDL.ActividadCorrelativo(actividad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EditarCausasDirectas")]
        public CausasDirectas EditarCausasDirectas(CausasDirectas causasDirectas)
        {
            try
            {
                return identificacionDL.EditarCausasDirectas(causasDirectas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        [HttpPost]
        [ActionName("EliminarCausasDirectas")]
        public CausasDirectas EliminarCausasDirectas(CausasDirectas causasDirectas)
        {
            try
            {
                return identificacionDL.EliminarCausasDirectas(causasDirectas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarCausasDirectas")]
        public List<CausasDirectas> ListarCausasDirectas(CausasDirectas causasDirectas)
        {
            try
            {
                return identificacionDL.ListarCausasDirectas(causasDirectas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarCausasIndirectas")]
        public CausasIndirectas InsertarCausasIndirectas(CausasIndirectas causasIndirectas)
        {
            try
            {
                return identificacionDL.InsertarCausasIndirectas(causasIndirectas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EditarCausasIndirectas")]
        public CausasIndirectas EditarCausasIndirectas(CausasIndirectas causasIndirectas)
        {
            try
            {
                return identificacionDL.EditarCausasIndirectas(causasIndirectas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EliminarCausasIndirectas")]
        public CausasIndirectas EliminarCausasIndirectas(CausasIndirectas causasIndirectas)
        {
            try
            {
                return identificacionDL.EliminarCausasIndirectas(causasIndirectas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarCausasIndirectas")]
        public List<CausasIndirectas> ListarCausasIndirectas(CausasIndirectas causasIndirectas)
        {
            try
            {
                return identificacionDL.ListarCausasIndirectas(causasIndirectas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarEfectoDirecto")]
        public EfectoDirecto InsertarEfectoDirecto(EfectoDirecto efectoDirecto)
        {
            try
            {
                return identificacionDL.InsertarEfectoDirecto(efectoDirecto);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EditarEfectoDirecto")]

        public EfectoDirecto EditarEfectoDirecto(EfectoDirecto efectoDirecto)
        {
            try
            {
                return identificacionDL.EditarEfectoDirecto(efectoDirecto);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EliminarEfectoDirecto")]
        public EfectoDirecto EliminarEfectoDirecto(EfectoDirecto efectoDirecto)
        {
            try
            {
                return identificacionDL.EliminarEfectoDirecto(efectoDirecto);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarEfectoDirecto")]
        public List<EfectoDirecto> ListarEfectoDirecto(EfectoDirecto efectoDirecto)
        {
            try
            {
                return identificacionDL.ListarEfectoDirecto(efectoDirecto);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarEfectoIndirecto")]
        public EfectoIndirecto InsertarEfectoIndirecto(EfectoIndirecto efectoIndirecto)
        {
            try
            {
                return identificacionDL.InsertarEfectoIndirecto(efectoIndirecto);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EditarEfectoIndirecto")]
        public EfectoIndirecto EditarEfectoIndirecto(EfectoIndirecto efectoIndirecto)
        {
            try
            {
                return identificacionDL.EditarEfectoIndirecto(efectoIndirecto);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EliminarIndirecto")]
        public EfectoIndirecto EliminarIndirecto(EfectoIndirecto efectoIndirecto)
        {
            try
            {
                return identificacionDL.EliminarIndirecto(efectoIndirecto);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarEfectoIndirecto")]
        public List<EfectoIndirecto> ListarEfectoIndirecto(EfectoIndirecto efectoIndirecto)
        {
            try
            {
                return identificacionDL.ListarEfectoIndirecto(efectoIndirecto);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarIndicador")]
        public Indicadores InsertarIndicador(Indicadores indicadores)
        {
            try
            {
                return identificacionDL.InsertarIndicador(indicadores);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EditarIndicador")]
        public Indicadores EditarIndicador(Indicadores indicadores)
        {
            try
            {
                return identificacionDL.EditarIndicador(indicadores);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("EliminarIndicador")]
        public Indicadores EliminarIndicador(Indicadores indicadores)
        {
            try
            {
                return identificacionDL.EliminarIndicador(indicadores);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        [HttpPost]
        [ActionName("InsertarCompDescrip")]
        public Componente InsertarCompDescrip(Componente componente)
        {
            try
            {
                return identificacionDL.InsertarCompDescrip(componente);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarComponentesSelect")]
        public List<Componente> ListarComponentesSelect(Componente componente)
        {
            try
            {
                return identificacionDL.ListarComponentesSelect(componente);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarIndicadoresPaginado")]
        public List<Indicadores> ListarIndicadoresPaginado(Indicadores indicadores)
        {
            try
            {
                return identificacionDL.ListarIndicadoresPaginado(indicadores);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

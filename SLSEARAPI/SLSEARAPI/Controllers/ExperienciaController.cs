using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace SLSEARAPI.Controllers
{
    public class ExperienciaController : ArchibobaseExperienciaController
    {
        ExperienciaDL experienciaDL;
        private Exception ex;

        public ExperienciaController()
        {

            experienciaDL = new ExperienciaDL();

        }


        [HttpPost]
        [ActionName("InsertarExperiencia")]
        public async Task<Experiencia> SubirArchivo()
        {
            try
            {
                Experiencia experiencia = new Experiencia();
                string path = HttpContext.Current.Request.Params["path"];
                string extension = ".pdf";//HttpContext.Current.Request.Params[".pdf"];

                var ruta = await ValidarSubirArchivosAsync(Request.Content.IsMimeMultipartContent(), path, extension);
                //var mensaje2 = ruta.mensaje;

                experiencia.vNombreEntidad = HttpContext.Current.Request.Params["vNombreEntidad"];
                experiencia.vCargoServicio = HttpContext.Current.Request.Params["vCargoServicio"];
                experiencia.dFechaInicio =   HttpContext.Current.Request.Params["dFechaInicio"];
                experiencia.dFechaFin = HttpContext.Current.Request.Params["dFechaFin"];
                experiencia.vRutaArchivoConstancia = HttpContext.Current.Request.Params["vRutaArchivoConstancia"];


                


                return experienciaDL.InsertarExperiencia(experiencia);
            }
            catch (Exception ex)
            {

                return new Experiencia
                {

                    mensaje = ex.Message,
                    validation = false
                };
            }
        }
    }
}

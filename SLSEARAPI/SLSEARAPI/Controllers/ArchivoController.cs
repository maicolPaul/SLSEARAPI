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
    public class ArchivoController : ArchivosbaseController
    {

        ArchivoDL archivoDL;
        private Exception ex;

        public ArchivoController()
        {

            archivoDL = new ArchivoDL();
 
        }


        [HttpPost]
        [ActionName("InsertarArchivo")]
        public async Task<Archivo> SubirArchivo()
        {
            try
            {
                Archivo archivo = new Archivo();
                string path = HttpContext.Current.Request.Params["path"];
                string extension = ".pdf";//HttpContext.Current.Request.Params[".pdf"];

                var ruta = await ValidarSubirArchivosAsync(Request.Content.IsMimeMultipartContent(), path, extension);
                //var mensaje2 = ruta.mensaje;

                    archivo.iCodArchivos= Convert.ToInt32(HttpContext.Current.Request.Params["iCodArchivos"]);
                    archivo.icodExtensionista= Convert.ToInt32(HttpContext.Current.Request.Params["icodExtensionista"]);
                    archivo.iCodNombreArchivo = Convert.ToInt32(HttpContext.Current.Request.Params["iCodNombreArchivo"]);
                    archivo.vRutaArchivo = HttpContext.Current.Request.Params["vRutaArchivo"];
                 
                    
                return archivoDL.InsertarArchivo(archivo);
            }
            catch (Exception ex)
            {

                return new Archivo
                {

                    mensaje = ex.Message,
                    validation = false
                };
            }
        }

    }
}

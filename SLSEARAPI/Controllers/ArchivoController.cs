using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
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
                archivo.vRutaArchivo = HttpContext.Current.Request.Params["vRutaArchivo"];
                string path = HttpContext.Current.Request.Params["path"];
                string extension = ".pdf";//HttpContext.Current.Request.Params[".pdf"];

                var ruta = await ValidarSubirArchivosAsync(Request.Content.IsMimeMultipartContent(), path, extension);
                //var mensaje2 = ruta.mensaje;

                    archivo.iCodArchivos= Convert.ToInt32(HttpContext.Current.Request.Params["iCodArchivos"]);
                    archivo.icodExtensionista= Convert.ToInt32(HttpContext.Current.Request.Params["icodExtensionista"]);
                    archivo.iCodNombreArchivo = Convert.ToInt32(HttpContext.Current.Request.Params["iCodNombreArchivo"]);
                    archivo.vRutaArchivo = ruta.fileNames[0];                 
                    
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

        [HttpPost]
        [ActionName("ListarArchivo")]
        public List<Archivo> ListarArchivo(Archivo entidad)
        {
            try
            {
                return archivoDL.ListarArchivo(entidad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public HttpResponseMessage Post()
        {
            HttpResponseMessage result = null;
            var httpRequest = HttpContext.Current.Request;
            if (httpRequest.Files.Count > 0)
            {
                var docfiles = new List<string>();
                foreach (string file in httpRequest.Files)
                {
                    var postedFile = httpRequest.Files[file];
                    var filePath = HttpContext.Current.Server.MapPath("~/" + postedFile.FileName);
                    postedFile.SaveAs(filePath);
                    docfiles.Add(filePath);
                }
                result = Request.CreateResponse(HttpStatusCode.Created, docfiles);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.BadRequest);
            }
            return result;
        }

        [ActionName("DescargarArchivoFile")]
        [HttpPost]
        public HttpResponseMessage DescargarArchivoFile([FromBody]Documento archivo)
        {
            var x = FTPaBase64(archivo.path);
            var stream = new MemoryStream(x.file);
            var result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue(x.mineType);
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = x.path
            };
            return result;
        }

        [ActionName("EliminarArchivo")]
        [HttpPost]
        public Archivo EliminarArchivo(Archivo entidad)
        {
            try
            {
                return archivoDL.EliminarArchivo(entidad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        //[HttpPost]
        //public ActionResult UploadFiles()
        //{
        //    string path = Server.MapPath("~/Content/Upload/");
        //    HttpFileCollectionBase files = htt;
        //    for (int i = 0; i < files.Count; i++)
        //    {
        //        HttpPostedFileBase file = files[i];
        //        file.SaveAs(path + file.FileName);
        //    }
        //    return Json(files.Count + " Files Uploaded!");
        //}
    }
}

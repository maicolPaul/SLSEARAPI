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
    public class NombreArchivoController : ApiController
    {
        NombreArchivoDL nombreArchivoDL;

        private Exception ex;
        public NombreArchivoController()
        {

            nombreArchivoDL = new NombreArchivoDL();
        }

        [HttpPost]
        [ActionName("ListarNombreArchivo")]

        public List<NombreArchivo> ListarNombreArchivo(NombreArchivo NombreArchivo)
        {
            try
            {
                return nombreArchivoDL.ListarNombreArchivo(NombreArchivo);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

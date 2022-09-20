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
    public class ExtensionistaController : ApiController
    {

        ExtensionistaDL extensionistaDL;

        private Exception ex;
        public ExtensionistaController()
        {

            extensionistaDL = new ExtensionistaDL();
        }

        [HttpPost]
        [ActionName("InsertarExtensionista")]

        public Extensionista InsertarDocumento(Extensionista Extensionista)
        {
            try
            {
                return extensionistaDL.InsertarExtensionista(Extensionista);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

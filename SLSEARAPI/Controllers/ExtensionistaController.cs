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

        [HttpPost]
        [ActionName("ActualizarPropuestaExtensionista")]
        public Extensionista ActualizarPropuestaExtensionista(Extensionista entidad)
        {
            try
            {
                return extensionistaDL.ActualizarPropuestaExtensionista(entidad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        [HttpPost]
        [ActionName("Login")]
        public Extensionista Login(Extensionista entidad)
        {
            try
            {
                return extensionistaDL.Login(entidad);
            }
            catch (Exception)
            {
                throw;
            }
        }

        [HttpPost]
        [ActionName("ListarExtensionistaPorCodigo")]
        public Extensionista ListarExtensionistaPorCodigo(Extensionista entidad)
        {
            try
            {
                return extensionistaDL.ListarExtensionistaPorCodigo(entidad);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

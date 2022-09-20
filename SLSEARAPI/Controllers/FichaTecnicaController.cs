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
    public class FichaTecnicaController : ApiController
    {
        FichaTecnicaDL fichaTecnicaDL;
        public FichaTecnicaController()
        {
            fichaTecnicaDL = new FichaTecnicaDL();
        }
        [HttpPost]
        [ActionName("ListarFichaTecnica")]
        public List<FichaTecnica> ListarFichaTecnica(FichaTecnica fichaTecnicapar)
        {
            try
            {
                return fichaTecnicaDL.ListarFichaTecnica(fichaTecnicapar);
            }
            catch (Exception)
            {
                throw;
            }
        }
        [HttpPost]
        [ActionName("InsertarFichaTecnica")]
        public FichaTecnica InsertarFichaTecnica(FichaTecnica fichaTecnica)
        {
            try
            {
                return fichaTecnicaDL.InsertarFichaTecnica(fichaTecnica);
            }
            catch (Exception)
            {
                throw;
            }
        }

        [HttpPost]
        [ActionName("ListarTipoProveedor")]
        public List<TipoProveedor> ListarTipoProveedor()
        {
            try
            {
                return fichaTecnicaDL.ListarTipoProveedor();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarSector")]
        public List<Sector> ListarSector()
        {
            try
            {
                return fichaTecnicaDL.ListarSector();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarCadenaProductivaAgraria")]
        public List<CadenaProductivaAgraria> ListarCadenaProductivaAgraria()
        {
            try
            {
                return fichaTecnicaDL.ListarCadenaProductivaAgraria();
            }
            catch (Exception ex)
            {
                throw ex;
            }            
        }

        [HttpPost]
        [ActionName("ListarLineaPrioritaria")]
        public List<LineaPrioritaria> ListarLineaPrioritaria(LineaPrioritaria lineaPrioritaria)
        {
            try
            {
                return fichaTecnicaDL.ListarLineaPrioritaria(lineaPrioritaria);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarTipoOrganizacion")]
        public List<TipoOrganizacion> ListarTipoOrganizacion()
        {
            try
            {
                return fichaTecnicaDL.ListarTipoOrganizacion();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

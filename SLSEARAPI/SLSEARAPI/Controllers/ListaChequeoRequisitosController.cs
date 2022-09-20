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
    public class ListaChequeoRequisitosController : ApiController
    {
        ListaChequeoRequisitosDL listaChequeoRequisitosDL;

        private Exception ex;
        public ListaChequeoRequisitosController()
        {

            listaChequeoRequisitosDL = new ListaChequeoRequisitosDL();
        }

        [HttpPost]
        [ActionName("InsertarListaChequeoRequisitos")]

        public ListaChequeoRequisitos InsertarListaChequeoRequisitos(ListaChequeoRequisitos ListaChequeRequisitos)
        {
            try
            {
                return listaChequeoRequisitosDL.InsertarListaChequeoRequisitos(ListaChequeRequisitos);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

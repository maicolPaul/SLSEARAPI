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
    public class FormacionAcademicaController : ApiController
    {
        FormacionAcademicaDL formacionAcademicaDL;

        private Exception ex;
        public FormacionAcademicaController()
        {

            formacionAcademicaDL = new FormacionAcademicaDL();
        }

        [HttpPost]
        [ActionName("InsertarFormacionAcademica")]

        public FormacionAcademica InsertarFormacionAcademica(FormacionAcademica FormacionAcademica)
        {
            try
            {
                return formacionAcademicaDL.InsertarFormacionAcademica(FormacionAcademica);
            }
            catch (Exception)
            {
                throw;
            }
        }
        [HttpPost]
        [ActionName("ListarFormacionAcademica")]
        public List<FormacionAcademica> ListarFormacionAcademica(FormacionAcademica entidad)
        {
            try
            {
                return formacionAcademicaDL.ListarFormacionAcademica(entidad);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

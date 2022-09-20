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
    public class CurriculumVitaeController : ApiController
    {
        CurriculumVitaeDL curriculumVitaeDL;
        private Exception ex;

        public CurriculumVitaeController()
        {

            curriculumVitaeDL = new CurriculumVitaeDL();

        }


        [HttpPost]
        [ActionName("InsertarCurriculumVitae")]

        public CurriculumVitae InsertarCurriculumVitae(CurriculumVitae CurriculumVitae)
        {
            try
            {
                return curriculumVitaeDL.InsertarCurriculumVitae(CurriculumVitae);
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}

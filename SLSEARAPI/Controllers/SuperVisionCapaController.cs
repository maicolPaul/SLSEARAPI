﻿using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace SLSEARAPI.Controllers
{
    public class SuperVisionCapaController : ApiController
    {
        SuperVisionCabCapDL superVisionCabCapDL;

        public SuperVisionCapaController()
        {
            superVisionCabCapDL = new SuperVisionCabCapDL();
        }

        [HttpPost]
        [ActionName("InsertarSuperVisionCabCap")]
        public SupervisionCapCab InsertarSuperVisionCabCap(SupervisionCapCab entidad)
        {
            try
            {
                return superVisionCabCapDL.InsertarSuperVisionCabCap(entidad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarSuperVisionDetCap")]
        public SupervisionCapDet InsertarSuperVisionDetCap(SupervisionCapDet entidad)
        {
            try
            {
                return superVisionCabCapDL.InsertarSuperVisionDetCap(entidad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarRubros")]
        public List<Rubro> ListarRubros()
        {
            try
            {
                return superVisionCabCapDL.ListarRubros();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarCriterio")]
        public List<Criterio> ListarCriterio(Criterio criterio)
        {
            try
            {
                return superVisionCabCapDL.ListarCriterio(criterio);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarCalificacion")]
        public List<Calificacion> ListarCalificacion()
        {
            try
            {
                return superVisionCabCapDL.ListarCalificacion();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
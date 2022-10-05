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
    public class PlanCapacitacionController : ApiController
    {
        PlanCapacitacionDL capacitacionDL;
        public PlanCapacitacionController()
        {
            capacitacionDL = new PlanCapacitacionDL();
        }

        [HttpPost]
        [ActionName("ListarPlanCapacitacion")]
        public List<PlanCapacitacion> ListarPlanCapacitacion(PlanCapacitacion planCapacitacion)
        {
            try
            {
                return capacitacionDL.ListarPlanCapacitacion(planCapacitacion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarPlanSesion")]
        public List<PlanSesion> ListarPlanSesion(PlanSesion planSesion)
        {
            try
            {
                return capacitacionDL.ListarPlanSesion(planSesion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarPlanCapacitacion")]
        public PlanCapacitacion InsertarPlanCapacitacion(PlanCapacitacion planCapacitacion)
        {
            try
            {
                return capacitacionDL.InsertarPlanCapacitacion(planCapacitacion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarPlanSesion")]
        public PlanSesion InsertarPlanSesion(PlanSesion planSesion)
        {
            try
            {
                return capacitacionDL.InsertarPlanSesion(planSesion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

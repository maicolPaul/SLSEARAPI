using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;


namespace SLSEARAPI.DataLayer
{
    public class PlanCapacitacionDL
    {
        public List<PlanCapacitacion> ListarPlanCapacitacion(PlanCapacitacion planCapacitacion)
        {
            List<PlanCapacitacion> lista = new List<PlanCapacitacion>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarPlanCapcaticacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", planCapacitacion.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", planCapacitacion.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", planCapacitacion.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", planCapacitacion.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodActividad", planCapacitacion.iCodActividad);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    planCapacitacion = new PlanCapacitacion();

                                    planCapacitacion.iCodPlanCap = dr.GetInt32(dr.GetOrdinal("iCodPlanCap"));
                                    planCapacitacion.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    planCapacitacion.vModuloTema = dr.GetString(dr.GetOrdinal("vModuloTema"));
                                    planCapacitacion.vObjetivo = dr.GetString(dr.GetOrdinal("vObjetivo"));
                                    planCapacitacion.iMeta = dr.GetInt32(dr.GetOrdinal("iMeta"));
                                    planCapacitacion.iBeneficiario = dr.GetInt32(dr.GetOrdinal("iBeneficiario"));
                                    planCapacitacion.dFechaActividad = dr.GetString(dr.GetOrdinal("dFechaActividad"));
                                    planCapacitacion.iTotalTeoria = dr.GetInt32(dr.GetOrdinal("iTotalTeoria"));
                                    planCapacitacion.iTotalPractica = dr.GetInt32(dr.GetOrdinal("iTotalPractica"));
                                    planCapacitacion.bActivo = dr.GetBoolean(dr.GetOrdinal("bActivo"));
                                    planCapacitacion.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    planCapacitacion.iCodHito = dr.GetInt32(dr.GetOrdinal("iCodHito"));
                                    lista.Add(planCapacitacion);
                                }
                            }
                        }
                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }

        public List<PlanSesion> ListarPlanSesion(PlanSesion planSesion)
        {
            List<PlanSesion> lista = new List<PlanSesion>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarPlanSesion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", planSesion.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", planSesion.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", planSesion.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", planSesion.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodPlanCap", planSesion.iCodPlanCap);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    planSesion = new PlanSesion();

                                    planSesion.iCodPlanCap = dr.GetInt32(dr.GetOrdinal("iCodPlanCap"));
                                    planSesion.iCodPlanSesion = dr.GetInt32(dr.GetOrdinal("iCodPlanSesion"));
                                    planSesion.iDuracion = dr.GetInt32(dr.GetOrdinal("iDuracion"));
                                    planSesion.vTematica = dr.GetString(dr.GetOrdinal("vTematica"));
                                    planSesion.vDescripMetodologia = dr.GetString(dr.GetOrdinal("vDescripMetodologia"));
                                    planSesion.vMateriales = dr.GetString(dr.GetOrdinal("vMateriales"));
                                    planSesion.bActivo = dr.GetBoolean(dr.GetOrdinal("bActivo"));
                                    planSesion.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    lista.Add(planSesion);
                                }
                            }
                        }
                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }

        public PlanCapacitacion InsertarPlanCapacitacion(PlanCapacitacion planCapacitacion)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarPlanCapacitacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodPlanCap", planCapacitacion.iCodPlanCap);
                        command.Parameters.AddWithValue("@iCodActividad", planCapacitacion.iCodActividad);
                        command.Parameters.AddWithValue("@vModuloTema", planCapacitacion.vModuloTema);
                        command.Parameters.AddWithValue("@vObjetivo", planCapacitacion.vObjetivo);
                        command.Parameters.AddWithValue("@iMeta", planCapacitacion.iMeta);
                        command.Parameters.AddWithValue("@iBeneficiario", planCapacitacion.iBeneficiario);
                        command.Parameters.AddWithValue("@dFechaActividad", Convert.ToDateTime(planCapacitacion.dFechaActividad));
                        command.Parameters.AddWithValue("@iTotalTeoria", planCapacitacion.iTotalTeoria);
                        command.Parameters.AddWithValue("@iTotalPractica", planCapacitacion.iTotalPractica);
                        command.Parameters.AddWithValue("@iopcion", planCapacitacion.iopcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    planCapacitacion.iCodPlanCap = dr.GetInt32(dr.GetOrdinal("iCodPlanCap"));
                                    planCapacitacion.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                }
                            }
                        }
                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return planCapacitacion;
        }

        public PlanSesion InsertarPlanSesion(PlanSesion planSesion)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarPlanSesion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodPlanCap", planSesion.iCodPlanCap);
                        command.Parameters.AddWithValue("@iCodPlanSesion", planSesion.iCodPlanSesion);
                        command.Parameters.AddWithValue("@iDuracion", planSesion.iDuracion);
                        command.Parameters.AddWithValue("@vTematica", planSesion.vTematica);
                        command.Parameters.AddWithValue("@vDescripMetodologia", planSesion.vDescripMetodologia);
                        command.Parameters.AddWithValue("@vMateriales", planSesion.vMateriales);
                        command.Parameters.AddWithValue("@iopcion", planSesion.iopcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    planSesion.iCodPlanSesion = dr.GetInt32(dr.GetOrdinal("iCodPlanSesion"));
                                    planSesion.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                }
                            }
                        }
                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return planSesion;
        }

        public List<Componente> PA_Listar_ComponentesPorExtensionista(Identificacion identificacion)
        {
            List<Componente> lista = new List<Componente>();
            Componente componente = null;

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Listar_ComponentesPorExtensionista]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", identificacion.iCodExtensionista);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    componente = new Componente();

                                    componente.iCodComponente = dr.GetInt32(dr.GetOrdinal("iCodComponente"));
                                    componente.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    lista.Add(componente);
                                }
                            }
                        }
                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }

        public DataTable Listar_PlanCapa_Rpt(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_PlanCapa_Rpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", actividad.iCodExtensionista);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        da.Fill(dataTable);

                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dataTable;
        }

        public DataTable SP_Listar_PlanCapa_Rpt2(PlanCapacitacion planCapacitacion)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_PlanCapa_Rpt2]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodActividad", planCapacitacion.iCodActividad);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        da.Fill(dataTable);

                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dataTable;
        }

        public DataTable SP_Listar_PlanCapa_Rpt3(PlanSesion planSesion)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_PlanCap_Rpt3]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodPlanCap", planSesion.iCodPlanCap);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        da.Fill(dataTable);

                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dataTable;
        }
    }
}
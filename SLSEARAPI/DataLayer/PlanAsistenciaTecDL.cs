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
    public class PlanAsistenciaTecDL
    {
        public List<PlanAsistenciaTec> ListarPlanAsistenciaTec(PlanAsistenciaTec planAsistenciaTec)
        {
            List<PlanAsistenciaTec> lista = new List<PlanAsistenciaTec>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarPlanAsistenciaTec]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", planAsistenciaTec.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", planAsistenciaTec.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", planAsistenciaTec.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", planAsistenciaTec.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodActividad", planAsistenciaTec.iCodActividad);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    planAsistenciaTec = new PlanAsistenciaTec();

                                    planAsistenciaTec.iCodPlanAsistenciaTec = dr.GetInt32(dr.GetOrdinal("iCodPlanAsistenciaTec"));
                                    planAsistenciaTec.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    //planAsistenciaTec.vModuloTema = dr.GetString(dr.GetOrdinal("vModuloTema"));
                                    planAsistenciaTec.vObjetivo = dr.GetString(dr.GetOrdinal("vObjetivo"));
                                    planAsistenciaTec.vObjetivoCorta = dr.GetString(dr.GetOrdinal("vObjetivoCorta"));
                                    planAsistenciaTec.iMeta = dr.GetInt32(dr.GetOrdinal("iMeta"));
                                    planAsistenciaTec.iBeneficiario = dr.GetInt32(dr.GetOrdinal("iBeneficiario"));
                                    //planAsistenciaTec.dFechaActividad = dr.GetString(dr.GetOrdinal("dFechaActividad"));
                                    //planAsistenciaTec.dFechaActividadFin = dr.GetString(dr.GetOrdinal("dFechaActividadFin"));
                                    planAsistenciaTec.iTotalTeoria = dr.GetDecimal(dr.GetOrdinal("iTotalTeoria"));
                                    planAsistenciaTec.iTotalPractica = dr.GetDecimal(dr.GetOrdinal("iTotalPractica"));
                                    planAsistenciaTec.bActivo = dr.GetBoolean(dr.GetOrdinal("bActivo"));
                                    planAsistenciaTec.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    planAsistenciaTec.iCodHito = dr.GetInt32(dr.GetOrdinal("iCodHito"));
                                    lista.Add(planAsistenciaTec);
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

        public List<PlanAsistenciaTecDet> ListarPlanAsistenciaTecDet(PlanAsistenciaTecDet planAsistenciaTecDet)
        {
            List<PlanAsistenciaTecDet> lista = new List<PlanAsistenciaTecDet>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarPlanAsistenciaTecDet]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", planAsistenciaTecDet.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", planAsistenciaTecDet.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", planAsistenciaTecDet.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", planAsistenciaTecDet.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodPlanAsistenciaTec", planAsistenciaTecDet.iCodPlanAsistenciaTec);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    planAsistenciaTecDet = new PlanAsistenciaTecDet();

                                    planAsistenciaTecDet.iCodPlanAsistenciaTecDet = dr.GetInt32(dr.GetOrdinal("iCodPlanAsistenciaTecDet"));
                                    planAsistenciaTecDet.iCodPlanAsistenciaTec = dr.GetInt32(dr.GetOrdinal("iCodPlanAsistenciaTec"));
                                    planAsistenciaTecDet.iDuracion = dr.GetInt32(dr.GetOrdinal("iDuracion"));
                                   // planAsistenciaTecDet.vTematica = dr.GetString(dr.GetOrdinal("vTematica"));
                                    planAsistenciaTecDet.vDescripMetodologia = dr.GetString(dr.GetOrdinal("vDescripMetodologia"));
                                    planAsistenciaTecDet.vDescripMetodologiaCorta = dr.GetString(dr.GetOrdinal("vDescripMetodologiaCorta"));
                                    planAsistenciaTecDet.vMateriales = dr.GetString(dr.GetOrdinal("vMateriales"));
                                    planAsistenciaTecDet.bActivo = dr.GetBoolean(dr.GetOrdinal("bActivo"));
                                    planAsistenciaTecDet.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    lista.Add(planAsistenciaTecDet);
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

        public PlanAsistenciaTec InsertarPlanAsistenciaTec(PlanAsistenciaTec planAsistenciaTec)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarPlanAsistenciaTec]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodPlanAsistenciaTec", planAsistenciaTec.iCodPlanAsistenciaTec);
                        command.Parameters.AddWithValue("@iCodActividad", planAsistenciaTec.iCodActividad);
                        //command.Parameters.AddWithValue("@vModuloTema", planAsistenciaTec.vModuloTema);
                        command.Parameters.AddWithValue("@vObjetivo", planAsistenciaTec.vObjetivo);
                        command.Parameters.AddWithValue("@iMeta", planAsistenciaTec.iMeta);
                        //command.Parameters.AddWithValue("@iBeneficiario", planAsistenciaTec.iBeneficiario);
                        //command.Parameters.AddWithValue("@dFechaActividad", Convert.ToDateTime(planAsistenciaTec.dFechaActividad));
                        //command.Parameters.AddWithValue("@dFechaActividadFin", Convert.ToDateTime(planAsistenciaTec.dFechaActividadFin));
                        command.Parameters.AddWithValue("@iTotalTeoria", planAsistenciaTec.iTotalTeoria);
                        command.Parameters.AddWithValue("@iTotalPractica", planAsistenciaTec.iTotalPractica);
                        command.Parameters.AddWithValue("@iopcion", planAsistenciaTec.iopcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    planAsistenciaTec.iCodPlanAsistenciaTec = dr.GetInt32(dr.GetOrdinal("iCodPlanAsistenciaTec"));
                                    planAsistenciaTec.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return planAsistenciaTec;
        }

        public PlanAsistenciaTecDet InsertarPlanAsistenciaTecDet(PlanAsistenciaTecDet planAsistenciaTecDet)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarPlanAsistenciaTecDet]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodPlanAsistenciaTecDet", planAsistenciaTecDet.iCodPlanAsistenciaTecDet);
                        command.Parameters.AddWithValue("@iCodPlanAsistenciaTec", planAsistenciaTecDet.iCodPlanAsistenciaTec);
                        command.Parameters.AddWithValue("@iDuracion", planAsistenciaTecDet.iDuracion);
                        //command.Parameters.AddWithValue("@vTematica", planAsistenciaTecDet.vTematica);
                        command.Parameters.AddWithValue("@vDescripMetodologia", planAsistenciaTecDet.vDescripMetodologia);
                        command.Parameters.AddWithValue("@vMateriales", planAsistenciaTecDet.vMateriales);
                        command.Parameters.AddWithValue("@iopcion", planAsistenciaTecDet.iopcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    planAsistenciaTecDet.iCodPlanAsistenciaTecDet = dr.GetInt32(dr.GetOrdinal("iCodPlanAsistenciaTecDet"));
                                    planAsistenciaTecDet.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return planAsistenciaTecDet;
        }

        public DataTable SP_Listar_PlanAsistenciaTecnica_Rpt(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_PlanAsistenciaTecnica_Rpt]", conection))
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

        public DataTable SP_Listar_PlanAsistenciaTecnica_Rpt2(PlanAsistenciaTec planAsistenciaTec)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_PlanAsistenciaTecnica_Rpt2]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodActividad", planAsistenciaTec.iCodActividad);
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

        public DataTable SP_Listar_PlanAsistenciaTecnica_Rpt3(PlanAsistenciaTecDet planAsistenciaTecDet)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_PlanAsistenciaTecnica_Rpt3]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodPlanAsistenciaTec", planAsistenciaTecDet.iCodPlanAsistenciaTec);
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
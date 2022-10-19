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
    public class CostoDL
    {
        public List<Actividad> ListarActividad(Actividad actividad)
        {
            List<Actividad> lista = new List<Actividad>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarActividadesPorExtensionista_Costo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", actividad.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", actividad.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", actividad.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", actividad.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodExtensionista", actividad.iCodExtensionista);
                        command.Parameters.AddWithValue("@iCodcomponente", actividad.iCodIdentificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    actividad = new Actividad();

                                    actividad.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    actividad.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    actividad.vActividad = dr.GetString(dr.GetOrdinal("vActividad"));
                                    actividad.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    actividad.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    actividad.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    actividad.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
                                    actividad.nTipoActividad = dr.GetInt32(dr.GetOrdinal("nTipoActividad"));
                                    actividad.bActivo = dr.GetBoolean(dr.GetOrdinal("bActivo"));
                                    actividad.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    lista.Add(actividad);
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

        public List<Costo> ListarCosto(Costo costo)
        {
            List<Costo> lista = new List<Costo>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", costo.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", costo.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", costo.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", costo.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodActividad", costo.iCodActividad);
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    costo = new Costo();

                                    costo.iCodCosto = dr.GetInt32(dr.GetOrdinal("iCodCosto"));
                                    //costo.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    //costo.iCodComponente = dr.GetInt32(dr.GetOrdinal("iCodComponente"));
                                    costo.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    costo.iTipoMatServ = dr.GetInt32(dr.GetOrdinal("iTipoMatServ"));
                                    costo.TipoMatServ = dr.GetString(dr.GetOrdinal("TipoMatServ"));
                                    costo.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    costo.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    costo.iCantidad = dr.GetInt32(dr.GetOrdinal("iCantidad"));
                                    costo.dCostoUnitario = dr.GetDecimal(dr.GetOrdinal("dCostoUnitario"));
                                    costo.dFecha = dr.GetString(dr.GetOrdinal("dFecha"));
                                    costo.Estado = dr.GetString(dr.GetOrdinal("Estado"));
                                    costo.bActivo = dr.GetBoolean(dr.GetOrdinal("bActivo"));
                                    costo.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    lista.Add(costo);
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

        public Costo InsertarCosto(Costo costo)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodCosto", costo.iCodCosto);
                        //command.Parameters.AddWithValue("@iCodIdentificacion", costo.iCodIdentificacion);
                        //command.Parameters.AddWithValue("@iCodComponente", costo.iCodComponente);
                        command.Parameters.AddWithValue("@iCodActividad", costo.iCodActividad);
                        command.Parameters.AddWithValue("@iTipoMatServ", costo.iTipoMatServ);
                        command.Parameters.AddWithValue("@vDescripcion", costo.vDescripcion);
                        command.Parameters.AddWithValue("@vUnidadMedida", costo.vUnidadMedida);
                        command.Parameters.AddWithValue("@iCantidad", costo.iCantidad);
                        command.Parameters.AddWithValue("@dCostoUnitario", costo.dCostoUnitario);
                        command.Parameters.AddWithValue("@dFecha", costo.dFecha);
                        command.Parameters.AddWithValue("@iopcion", costo.iopcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    costo.iCodCosto = dr.GetInt32(dr.GetOrdinal("iCodCosto"));
                                    costo.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return costo;
        }

        public DataTable ListarComponentesRpt(Cronograma cronograma)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarComponentesCabeceraCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", cronograma.iCodExtensionista);
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

        public DataTable ListarActividadesPorComponente(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarActividadesPorComponente2Costo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodComponente", actividad.iCodIdentificacion);
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

        public DataTable ListarCostosPorActividad(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCostosPorActividad_RptCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodActividad", actividad.iCodActividad);
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

        public DataTable ListarCostosResumenPorComponente(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCostosResumenPorComponente_RptCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iTipoMatServ", actividad.iopcion);
                        command.Parameters.AddWithValue("@iCodComponente", actividad.iCodIdentificacion);
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

        public DataTable ListarCostosResumenGeneral(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCostosResumenGeneral_RptCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iTipoMatServ", actividad.iopcion);
                        //command.Parameters.AddWithValue("@iCodComponente", actividad.iCodIdentificacion);
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

    }
}
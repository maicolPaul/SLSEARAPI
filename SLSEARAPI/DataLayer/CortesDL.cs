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
    public class CortesDL
    {
        public CortesCabecera InsertarCorteCabecera(CortesCabecera cortesCabecera)
        {
			try
			{
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarCortesCabecera]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodFichaTecnica", cortesCabecera.iCodFichaTecnica);
                        command.Parameters.AddWithValue("@dFechaInicioReal", cortesCabecera.dFechaInicioReal);
                        command.Parameters.AddWithValue("@dFechaFinReal", cortesCabecera.dFechaFinReal);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    cortesCabecera.iCodCorte = dr.GetInt32(dr.GetOrdinal("iCodCorte"));
                                    cortesCabecera.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return cortesCabecera;
        }
        public CortesDetalle InsertarCorteDetalle(CortesDetalle cortesDetalle)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarCortesDetalle]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodCorte", cortesDetalle.iCodCorte);
                        command.Parameters.AddWithValue("@idias", cortesDetalle.idias);
                       
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    cortesDetalle.iCodCorteDetalle = dr.GetInt32(dr.GetOrdinal("iCodCorteDetalle"));
                                    cortesDetalle.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return cortesDetalle;
        }
        public List<CortesDetalle> ListarCortesDetalle(CortesDetalle cortesDetalle)
        {
            List<CortesDetalle> lista = new List<CortesDetalle>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCortesDetalle]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodCorte", cortesDetalle.iCodCorte);
                        command.Parameters.AddWithValue("@piPageSize", cortesDetalle.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", cortesDetalle.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", cortesDetalle.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", cortesDetalle.pvSortOrder);


                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while(dr.Read())
                                {
                                    cortesDetalle = new CortesDetalle();
                                    cortesDetalle.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    cortesDetalle.totalPaginas = dr.GetInt32(dr.GetOrdinal("iPageCount"));
                                    cortesDetalle.paginaActual = dr.GetInt32(dr.GetOrdinal("iCurrentPage"));
                                    cortesDetalle.iCodCorteDetalle = dr.GetInt32(dr.GetOrdinal("iCodCorteDetalle"));
                                    cortesDetalle.iCodCorte = dr.GetInt32(dr.GetOrdinal("iCodCorte"));
                                    cortesDetalle.idias = dr.GetInt32(dr.GetOrdinal("idias"));
                                    cortesDetalle.Entregable = dr.GetString(dr.GetOrdinal("Entregable"));
                                    lista.Add(cortesDetalle);
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
        public CortesDetalle EliminarCorteDetalle(CortesDetalle cortesDetalle)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EliminarCorteDetalle]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCorteDetalle", cortesDetalle.iCodCorteDetalle);
           
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    cortesDetalle.iCodCorteDetalle = dr.GetInt32(dr.GetOrdinal("iCodCorteDetalle"));
                                    cortesDetalle.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return cortesDetalle;
        }
        public CortesCabecera ObtenerCorteCabecera(CortesCabecera cortesCabecera)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ObtenerCorteCabecera]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodFichaTecnica", cortesCabecera.iCodFichaTecnica);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    cortesCabecera.iCodCorte = dr.GetInt32(dr.GetOrdinal("iCodCorte"));
                                    cortesCabecera.dFechaInicioReal = dr.GetString(dr.GetOrdinal("dFechaInicioReal"));
                                    cortesCabecera.dFechaFinReal = dr.GetString(dr.GetOrdinal("dFechaFinReal"));                                                                       
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

            return cortesCabecera;
        }
    }
}
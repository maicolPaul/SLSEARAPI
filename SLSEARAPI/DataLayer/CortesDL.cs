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
                                    cortesDetalle.iCodCorte = dr.GetInt32(dr.GetOrdinal("iCodCorteDetalle"));
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
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while(dr.Read())
                                {
                                    cortesDetalle = new CortesDetalle();
                                    cortesDetalle.iCodCorteDetalle = dr.GetInt32(dr.GetOrdinal("iCodCorteDetalle"));
                                    cortesDetalle.iCodCorte = dr.GetInt32(dr.GetOrdinal("iCodCorte"));
                                    cortesDetalle.idias = dr.GetInt32(dr.GetOrdinal("idias"));
                                    cortesDetalle.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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
    }
}
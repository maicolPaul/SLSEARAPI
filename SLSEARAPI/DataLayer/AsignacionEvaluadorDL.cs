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
    public class AsignacionEvaluadorDL
    {
        public List<Identificacion> ListarSear(Identificacion identificacion)
        {
            List<Identificacion> lista = new List<Identificacion>();
            
			try
			{
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarSears]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", identificacion.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", identificacion.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", identificacion.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", identificacion.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodUbigeoT1", identificacion.iCodUbigeoT1);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    identificacion = new Identificacion();
                                    identificacion.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    identificacion.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    identificacion.vNombreSearT1 = dr.GetString(dr.GetOrdinal("vNombreSearT1"));
                                    identificacion.vEstado = dr.GetString(dr.GetOrdinal("ESTADO"));
                                    lista.Add(identificacion);
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
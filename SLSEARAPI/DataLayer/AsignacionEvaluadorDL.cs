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
        public ComiteIdentificacion AsignacionEvaluador(ComiteIdentificacion comiteIdentificacion)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Insertar_ComiteIdentificacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@vCodUbigeo", comiteIdentificacion.iCodUbigeoT1);
                        command.Parameters.AddWithValue("@iCodIdentificacion", comiteIdentificacion.iCodIdentificacion);                        

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    comiteIdentificacion = new ComiteIdentificacion();

                                    comiteIdentificacion.iCodComiteIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodComiteIdentificacion"));
                                    comiteIdentificacion.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));                                    
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
            return comiteIdentificacion;
        }
        public ComiteEvaluador EliminarComiteEvaluadorPorIdentificacion(ComiteEvaluador comiteEvaluador)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Eliminar_ComiteEvaluadorPorIdentificacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        //command.Parameters.AddWithValue("@vCodUbigeo", comiteIdentificacion.iCodUbigeoT1);
                        command.Parameters.AddWithValue("@iCodIdentificacion", comiteEvaluador.iCodIdentificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    comiteEvaluador = new ComiteEvaluador();

                                    comiteEvaluador.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    comiteEvaluador.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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
            return comiteEvaluador;
        }
        public List<ComiteEvaluador> ListarComiteEvaluadorPorIdentificacion(ComiteEvaluador comiteEvaluador)
        {
            List<ComiteEvaluador> lista = new List<ComiteEvaluador>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Listar_ComiteEvaluadorPorIdentificacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;                        
                        command.Parameters.AddWithValue("@iCodIdentificacion", comiteEvaluador.iCodIdentificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    comiteEvaluador = new ComiteEvaluador();

                                    comiteEvaluador.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    comiteEvaluador.iCodComiteEvaluador = dr.GetInt32(dr.GetOrdinal("iCodComiteEvaluador"));
                                    comiteEvaluador.vNombres = dr.GetString(dr.GetOrdinal("vNombres"));
                                    comiteEvaluador.vApellidoPat = dr.GetString(dr.GetOrdinal("vApellidoPat"));
                                    comiteEvaluador.vApellidoMat = dr.GetString(dr.GetOrdinal("vApellidoMat"));                                    
                                    comiteEvaluador.vCorreo = dr.GetString(dr.GetOrdinal("vCorreo"));
                                    lista.Add(comiteEvaluador);
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
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
    public class FichaEvaluacionDL
    {
        public List<ComiteIdentificacion> ListarComiteIdentificacion(ComiteIdentificacion comiteIdentificacion)
        {
            List<ComiteIdentificacion> lista = new List<ComiteIdentificacion>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_ListarComiteIdentificacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", comiteIdentificacion.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", comiteIdentificacion.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", comiteIdentificacion.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", comiteIdentificacion.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodComiteEvaluador", comiteIdentificacion.iCodComiteEvaluador);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                FichaTecnica fichatecnica;
                                while (dr.Read())
                                {
                                    comiteIdentificacion = new ComiteIdentificacion();
                                    comiteIdentificacion.iCodComiteIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodComiteIdentificacion"));
                                    comiteIdentificacion.iCodComiteEvaluador = dr.GetInt32(dr.GetOrdinal("iCodComiteEvaluador"));
                                    comiteIdentificacion.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    comiteIdentificacion.vNombreSearT1 = dr.GetString(dr.GetOrdinal("vNombreSearT1"));
                                    comiteIdentificacion.vDireccionT2 = dr.GetString(dr.GetOrdinal("vDireccionT2"));
                                    comiteIdentificacion.iCodUbigeoT1 = dr.GetString(dr.GetOrdinal("iCodUbigeoT1"));
                                    comiteIdentificacion.vNomDepartamento = dr.GetString(dr.GetOrdinal("vNomDepartamento"));
                                    comiteIdentificacion.vNomProvincia = dr.GetString(dr.GetOrdinal("vNomProvincia"));
                                    comiteIdentificacion.vNomDistrito = dr.GetString(dr.GetOrdinal("vNomDistrito"));
                                    comiteIdentificacion.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    //iRecordCount
                                    lista.Add(comiteIdentificacion);
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

        public List<FichaEvaluacion> ListarFichaEvaluacion(FichaEvaluacion fichaEvaluacion)
        {
            List<FichaEvaluacion> lista = new List<FichaEvaluacion>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_ListarFichaEvaluacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", fichaEvaluacion.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", fichaEvaluacion.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", fichaEvaluacion.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", fichaEvaluacion.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodIdentificacion", fichaEvaluacion.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodComiteEvaluador", fichaEvaluacion.iCodComiteEvaluador);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    fichaEvaluacion = new FichaEvaluacion();
                                    fichaEvaluacion.iCodFichaEvaluacion = dr.GetInt32(dr.GetOrdinal("iCodFichaEvaluacion"));
                                    fichaEvaluacion.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    fichaEvaluacion.iCodComiteEvaluador = dr.GetInt32(dr.GetOrdinal("iCodComiteEvaluador"));
                                    fichaEvaluacion.iCodCategoria = dr.GetInt32(dr.GetOrdinal("iCodCategoria"));
                                    fichaEvaluacion.vCategoria = dr.GetString(dr.GetOrdinal("vCategoria"));
                                    fichaEvaluacion.iCodCriterio = dr.GetInt32(dr.GetOrdinal("iCodCriterio"));
                                    fichaEvaluacion.vCriterio = dr.GetString(dr.GetOrdinal("vCriterio"));
                                    fichaEvaluacion.vCriterio1 = dr.GetString(dr.GetOrdinal("vCriterio1"));
                                    fichaEvaluacion.PuntajeMaximo = dr.GetDecimal(dr.GetOrdinal("PuntajeMaximo"));
                                    fichaEvaluacion.dPuntajeEvaluacion = dr.GetDecimal(dr.GetOrdinal("dPuntajeEvaluacion"));
                                    fichaEvaluacion.vJustificacion = dr.GetString(dr.GetOrdinal("vJustificacion")); 
                                    fichaEvaluacion.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    //iRecordCount
                                    lista.Add(fichaEvaluacion);
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

        public FichaEvaluacion InsertarFichaEvaluacion(FichaEvaluacion fichaEvaluacion)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Insertar_FichaEvaluacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodFichaEvaluacion", fichaEvaluacion.iCodFichaEvaluacion);
                        command.Parameters.AddWithValue("@iCodIdentificacion", fichaEvaluacion.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodComiteEvaluador", fichaEvaluacion.iCodComiteEvaluador);
                        command.Parameters.AddWithValue("@iCodCategoria", fichaEvaluacion.iCodCategoria);
                        command.Parameters.AddWithValue("@iCodCriterio", fichaEvaluacion.iCodCriterio);
                        command.Parameters.AddWithValue("@dPuntajeEvaluacion", fichaEvaluacion.dPuntajeEvaluacion);
                        command.Parameters.AddWithValue("@vJustificacion", fichaEvaluacion.vJustificacion);
                        command.Parameters.AddWithValue("@iOpcion", fichaEvaluacion.iOpcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    fichaEvaluacion.iCodFichaEvaluacion = dr.GetInt32(dr.GetOrdinal("iCodFichaEvaluacion"));
                                    fichaEvaluacion.vMensaje = dr.GetString(dr.GetOrdinal("vmensaje"));
                                }
                            }
                        }
                    }
                    conection.Close();
                }
                return fichaEvaluacion;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
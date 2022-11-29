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
    public class SuperVisionCabCapDL
    {
        public SupervisionCapCab InsertarSuperVisionCabCap(SupervisionCapCab entidad)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_InsertarSupervisionCapCab]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodIdentificacion", entidad.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodFichaTecnica", entidad.iCodFichaTecnica);
                        command.Parameters.AddWithValue("@iCodComponente", entidad.iCodComponente);
                        command.Parameters.AddWithValue("@iCodActividad", entidad.iCodActividad);
                        command.Parameters.AddWithValue("@vObservaciongeneral", entidad.vObservaciongeneral);
                        command.Parameters.AddWithValue("@vRecomendacion", entidad.vRecomendacion);
                        command.Parameters.AddWithValue("@vNombreSupervisor", entidad.vNombreSupervisor);
                        command.Parameters.AddWithValue("@vCargoSupervisor", entidad.vCargoSupervisor);
                        command.Parameters.AddWithValue("@vEntidadSupervisor", entidad.vEntidadSupervisor);
                        command.Parameters.AddWithValue("@dFechaSupervisor", entidad.dFechaSupervisor);
                        command.Parameters.AddWithValue("@iCodCalificacion", entidad.iCodCalificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {

                                    entidad.iCodSuperCab = dr.GetInt32(dr.GetOrdinal("iCodSuperCab"));

                                    entidad.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                }
                            }
                        }

                    }
                    conection.Close();
                }
                return entidad;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public SupervisionCapDet InsertarSuperVisionDetCap(SupervisionCapDet entidad)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_InsertarSuperVisionDetCapa]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodSuperCab", entidad.iCodSuperCab);
                        command.Parameters.AddWithValue("@iCodRubro", entidad.iCodRubro);
                        command.Parameters.AddWithValue("@vFundamento", entidad.vFundamento);
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {

                                    entidad.iCodSuperDet = dr.GetInt32(dr.GetOrdinal("iCodSuperDet"));

                                    entidad.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                }
                            }
                        }

                    }
                    conection.Close();
                }
                return entidad;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<Rubro> ListarRubros()
        {
            List<Rubro> lista = new List<Rubro>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_Listar_Rubros]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Rubro rubro;
                                while (dr.Read())
                                {
                                    rubro = new Rubro();
                                    rubro.iCodRubro = dr.GetInt32(dr.GetOrdinal("iCodRubro"));
                                    rubro.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    lista.Add(rubro);
                                }
                            }
                        }

                    }
                    conection.Close();
                }
                return lista;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<Criterio> ListarCriterio(Criterio criterio)
        {
            List<Criterio> lista=new List<Criterio>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_ListarCriterios]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", criterio.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", criterio.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", criterio.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", criterio.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodRubro", criterio.iCodRubro);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {                               
                                while (dr.Read())
                                {
                                    criterio = new Criterio();
                                    criterio.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    criterio.iPageCount = dr.GetInt32(dr.GetOrdinal("iPageCount"));
                                    criterio.iCurrentPage = dr.GetInt32(dr.GetOrdinal("iCurrentPage"));
                                    criterio.iCodCriterio = dr.GetInt32(dr.GetOrdinal("iCodCriterio"));
                                    criterio.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));              
                                    lista.Add(criterio);
                                }
                            }
                        }

                    }
                    conection.Close();
                }
                return lista;
            }
            catch (Exception)
            {
                throw;
            }         
        }
        public List<Calificacion>  ListarCalificacion()
        {
            List<Calificacion> lista = new List<Calificacion>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_ListarCalificacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
            
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Calificacion calificacion;
                                while (dr.Read())
                                {
                                    calificacion = new Calificacion();
                                    calificacion.iCodCalificacion = dr.GetInt32(dr.GetOrdinal("iCodCalificacion"));
                                    calificacion.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    lista.Add(calificacion);
                                }
                            }
                        }

                    }
                    conection.Close();
                }
                return lista;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
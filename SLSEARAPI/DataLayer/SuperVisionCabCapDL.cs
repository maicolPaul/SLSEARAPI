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
                        command.Parameters.AddWithValue("@iCodSuperCab", entidad.iCodSuperCab);
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
                        command.Parameters.AddWithValue("@iCodProductor", entidad.iCodProductor);

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
        public SupervisionCapDet2 InsertarSuperVisionDet2Cap(SupervisionCapDet2 entidad)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_InsertarSuperVisionDet2Capa]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodSuperCab", entidad.iCodSuperCab);
                        command.Parameters.AddWithValue("@iCodRubro", entidad.iCodRubro);
                        command.Parameters.AddWithValue("@iCodCriterio", entidad.iCodCrtierio);
                        command.Parameters.AddWithValue("@iCodCalificacion", entidad.iCodCalificacion);
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
        public SupervisionCapCab ObtenerSupervisionCapCab(SupervisionCapCab entidad)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_Obtener_Supervisio_CapCab]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodIdentificacion", entidad.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodFichaTecnica", entidad.iCodFichaTecnica);
                        command.Parameters.AddWithValue("@iCodComponente", entidad.iCodComponente);
                        command.Parameters.AddWithValue("@iCodActividad", entidad.iCodActividad);               
                        command.Parameters.AddWithValue("@iCodCalificacion", entidad.iCodCalificacion);
                        command.Parameters.AddWithValue("@iCodProductor", entidad.iCodProductor);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {

                                    entidad.iCodSuperCab = dr.GetInt32(dr.GetOrdinal("iCodSuperCab"));

                                    entidad.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    entidad.iCodFichaTecnica = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    entidad.vObservaciongeneral = dr.GetString(dr.GetOrdinal("vObservaciongeneral"));
                                    entidad.vRecomendacion = dr.GetString(dr.GetOrdinal("vRecomendacion"));
                                    entidad.vNombreSupervisor = dr.GetString(dr.GetOrdinal("vNombreSupervisor"));
                                    entidad.vCargoSupervisor = dr.GetString(dr.GetOrdinal("vCargoSupervisor"));
                                    entidad.vEntidadSupervisor = dr.GetString(dr.GetOrdinal("vEntidadSupervisor"));
                                    entidad.dFechaSupervisor = dr.GetString(dr.GetOrdinal("dFechaSupervisor"));


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
        public List<Rubro> ListarRubros(Rubro rubro)
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
                        command.Parameters.AddWithValue("@iCodSuperCab", rubro.iCodSuperCab);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {                           
                                while (dr.Read())
                                {
                                    rubro = new Rubro();
                                    rubro.iCodRubro = dr.GetInt32(dr.GetOrdinal("iCodRubro"));
                                    rubro.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    rubro.vFundamento = dr.GetString(dr.GetOrdinal("vFundamento"));
                                    rubro.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
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
                        command.Parameters.AddWithValue("@iCodSuperCab", criterio.iCodSuperCab);

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
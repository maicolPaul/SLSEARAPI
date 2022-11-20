using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;

namespace SLSEARAPI.DataLayer
{
    public class CronogramaDL
    {

        //public List<Componente> ListarComponentes(Cronograma cronograma)
        //{
        //    List<Componente> lista = new List<Componente>();

        //    try
        //    {
        //        using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
        //        {
        //            conection.Open();

        //            using (SqlCommand command = new SqlCommand("[PA_ListarComponentesCabecera]", conection))
        //            {
        //                command.CommandType = CommandType.StoredProcedure;
        //                command.Parameters.AddWithValue("@iCodExtensionista", cronograma.iCodExtensionista);


        //                using (SqlDataReader dr = command.ExecuteReader())
        //                {
        //                    if (dr.HasRows)
        //                    {
        //                        Componente componente;
        //                        while (dr.Read())
        //                        {
        //                            componente = new Componente();

        //                            componente.iCodComponente = dr.GetInt32(dr.GetOrdinal("iCodComponente"));
        //                            componente.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
        //                            componente.vDescComponente = dr.GetString(dr.GetOrdinal("vDescComponente"));
        //                            componente.vIndicador = dr.GetString(dr.GetOrdinal("vIndicador"));
        //                            componente.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));                                    
        //                            componente.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
        //                            componente.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
        //                            componente.nTipoComponente = dr.GetInt32(dr.GetOrdinal("nTipoComponente"));
                                    
        //                            lista.Add(componente);
        //                        }
        //                    }
        //                }
        //            }
        //            conection.Close();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }

        //    return lista;
        //}

        public DataTable ListarComponentesRpt(Cronograma cronograma)
        {
            //List<Componente> lista = new List<Componente>();
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarComponentesCabecera]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", cronograma.iCodExtensionista);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        // this will query your database and return the result to your datatable
                        da.Fill(dataTable);

                        //using (SqlDataReader dr = command.ExecuteReader())
                        //{
                        //    if (dr.HasRows)
                        //    {
                        //        Componente componente;
                        //        while (dr.Read())
                        //        {
                        //            componente = new Componente();

                        //            componente.iCodComponente = dr.GetInt32(dr.GetOrdinal("iCodComponente"));
                        //            componente.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                        //            componente.vDescComponente = dr.GetString(dr.GetOrdinal("vDescComponente"));
                        //            componente.vIndicador = dr.GetString(dr.GetOrdinal("vIndicador"));
                        //            componente.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                        //            componente.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                        //            componente.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
                        //            componente.nTipoComponente = dr.GetInt32(dr.GetOrdinal("nTipoComponente"));

                        //            lista.Add(componente);
                        //        }
                        //    }
                        //}
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

        public List<Cronograma> ListarConogramaFechaTipo(Cronograma cronograma)
        {
            List<Cronograma> lista = new List<Cronograma>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCronogramaFechasTipo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", cronograma.iCodExtensionista);
                        command.Parameters.AddWithValue("@nTipoActividad", cronograma.nTipoActividad);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    cronograma = new Cronograma();

                                    cronograma.iCodCronograma = dr.GetInt32(dr.GetOrdinal("iCodCronograma"));
                                    cronograma.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    cronograma.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    cronograma.vActividad = dr.GetString(dr.GetOrdinal("vActividad"));
                                    cronograma.vDescripcionActividad = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    cronograma.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    cronograma.iCantidad = dr.GetInt32(dr.GetOrdinal("iCantidad"));
                                    cronograma.iCodComponente = dr.GetInt32(dr.GetOrdinal("nTipoActividad"));
                                    cronograma.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    //cronograma.dFecha = dr.GetString(dr.GetOrdinal("dFecha"));
                                    cronograma.dfechacronograma = dr.GetDateTime(dr.GetOrdinal("dFecha"));
                                    lista.Add(cronograma);
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

        public List<Cronograma> ListarConogramaFecha(Cronograma cronograma)
        {
            List<Cronograma> lista = new List<Cronograma>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCronogramaFechas]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;                        
                        command.Parameters.AddWithValue("@iCodExtensionista", cronograma.iCodExtensionista);                        

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    cronograma = new Cronograma();
                                    
                                    cronograma.iCodCronograma = dr.GetInt32(dr.GetOrdinal("iCodCronograma"));
                                    cronograma.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    cronograma.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));                                    
                                    cronograma.vActividad = dr.GetString(dr.GetOrdinal("vActividad"));
                                    cronograma.vDescripcionActividad = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    cronograma.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    cronograma.iCantidad = dr.GetInt32(dr.GetOrdinal("iCantidad"));
                                    cronograma.iCodComponente = dr.GetInt32(dr.GetOrdinal("nTipoActividad"));
                                    cronograma.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    cronograma.dfechacronograma = dr.GetDateTime(dr.GetOrdinal("dFecha"));

                                    lista.Add(cronograma);
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
        public List<Cronograma> ListarCronograma(Cronograma cronograma)
        {
            List<Cronograma> lista = new List<Cronograma>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCronograma]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", cronograma.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", cronograma.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", cronograma.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", cronograma.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodExtensionista", cronograma.iCodExtensionista);
                        command.Parameters.AddWithValue("@iCodComponente", cronograma.iCodComponente);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {                                
                                while (dr.Read())
                                {
                                    cronograma = new Cronograma();
                                    cronograma.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    cronograma.totalPaginas = dr.GetInt32(dr.GetOrdinal("iPageCount"));
                                    cronograma.paginaActual = dr.GetInt32(dr.GetOrdinal("iCurrentPage"));
                                    cronograma.iCodCronograma = dr.GetInt32(dr.GetOrdinal("iCodCronograma"));
                                    cronograma.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    cronograma.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    cronograma.vDescripcionActividad = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    cronograma.vActividad = dr.GetString(dr.GetOrdinal("vActividad"));
                                    cronograma.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    cronograma.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    cronograma.iCantidad = dr.GetInt32(dr.GetOrdinal("iCantidad"));
                                    cronograma.iCodComponente = dr.GetInt32(dr.GetOrdinal("nTipoActividad"));
                                    cronograma.dFecha = dr.GetString(dr.GetOrdinal("dFecha"));

                                    lista.Add(cronograma);
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
        public List<Actividad> ListarActividades(Actividad actividad)
        {
            List<Actividad> lista = new List<Actividad>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarActividadesPorExtensionista2]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;                        
                        command.Parameters.AddWithValue("@iCodExtensionista", actividad.iCodExtensionista);
                        command.Parameters.AddWithValue("@nTipoActividad", actividad.nTipoActividad);


                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    actividad = new Actividad();                              
                                    actividad.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    actividad.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    actividad.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    actividad.vActividad = dr.GetString(dr.GetOrdinal("vActividad"));
                                    actividad.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    actividad.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    actividad.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    actividad.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));                                    

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
        public List<ComponenteCronograma> ListarComponentes(Identificacion indicadores)
        {
            List<ComponenteCronograma> lista = new List<ComponenteCronograma>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarComponentes]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCoExtensionista", indicadores.iCodExtensionista);
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                ComponenteCronograma componente;
                                while (dr.Read())
                                {
                                    componente = new ComponenteCronograma();
                                    componente.Codigo = dr.GetInt32(dr.GetOrdinal("Codigo"));
                                    componente.vDescComponente = dr.GetString(dr.GetOrdinal("vDescComponente"));

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

        public List<Actividad> ListaActividades(Identificacion identificacion)
        {
            List<Actividad> lista = new List<Actividad>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarActividadesPorExtensionista]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", identificacion.iCodExtensionista);
                        command.Parameters.AddWithValue("@tipoactividad", identificacion.iOpcion);
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Actividad componente;
                                while (dr.Read())
                                {
                                    componente = new Actividad();
                                    componente.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    componente.vActividad = dr.GetString(dr.GetOrdinal("vActividad"));
                                    componente.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    componente.resumen = dr.GetInt32(dr.GetOrdinal("resumen"));
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
        public Cronograma InsertarCronograma(Cronograma cronograma)
        {   
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarCronograma]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", cronograma.iCodExtensionista);
                        command.Parameters.AddWithValue("@iCodCronograma", cronograma.iCodCronograma);
                        command.Parameters.AddWithValue("@iCodComponente", cronograma.iCodComponente);
                        command.Parameters.AddWithValue("@iCodActividad", cronograma.iCodActividad);
                        command.Parameters.AddWithValue("@iCantidad", cronograma.iCantidad);
                        command.Parameters.AddWithValue("@dFecha", cronograma.dFecha);
                        command.Parameters.AddWithValue("@iopcion", cronograma.iopcion);
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                            
                                if (dr.Read())
                                {
                                    cronograma.iCodCronograma = dr.GetInt32(dr.GetOrdinal("iCodCronograma"));
                                    cronograma.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return cronograma;
        }

        public DataTable ListarActividadesPorComponente(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarActividadesPorComponente2]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodComponente", actividad.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodExtensionista", actividad.iCodExtensionista);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        da.Fill(dataTable);

                        //using (SqlDataReader dr = command.ExecuteReader())
                        //{
                        //    if (dr.HasRows)
                        //    {
                        //        while (dr.Read())
                        //        {
                        //            actividad = new Actividad();
                        //            actividad.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                        //            actividad.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                        //            actividad.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                        //            actividad.vActividad = dr.GetString(dr.GetOrdinal("vActividad"));
                        //            actividad.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                        //            actividad.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                        //            actividad.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                        //            actividad.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));

                        //            lista.Add(actividad);
                        //        }
                        //    }
                        //}
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
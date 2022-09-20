using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace SLSEARAPI.DataLayer
{
    public class ExperienciaDL
    {
        public List<Experiencia> ListarExperiencia(Experiencia entidad)
        {
            List<Experiencia> lista = new List<Experiencia>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_ListarExperiencia]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@piPageSize", entidad.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", entidad.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", entidad.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", entidad.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodCurriculumVitae", entidad.iCodCurriculumVitae);
                        command.Parameters.AddWithValue("@iTipoExperiencia", entidad.iTipoExperiencia);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Experiencia experiencia;
                                while (dr.Read())
                                {
                                    experiencia = new Experiencia();
                                    experiencia.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    experiencia.totalPaginas = dr.GetInt32(dr.GetOrdinal("iPageCount"));
                                    experiencia.paginaActual = dr.GetInt32(dr.GetOrdinal("iCurrentPage"));
                                    experiencia.iCodExperiencia = dr.GetInt32(dr.GetOrdinal("iCodExperiencia"));                                    
                                    experiencia.vNombreEntidad = dr.GetString(dr.GetOrdinal("vNombreEntidad"));
                                    experiencia.vCargoServicio = dr.GetString(dr.GetOrdinal("vCargoServicio"));
                                    experiencia.vActividades = dr.GetString(dr.GetOrdinal("vActividades"));
                                    experiencia.vProductoServicio = dr.GetString(dr.GetOrdinal("vProductoServicio"));
                                    experiencia.dFechaInicio = dr.GetString(dr.GetOrdinal("dFechaInicio"));
                                    experiencia.dFechaFin = dr.GetString(dr.GetOrdinal("dFechaFin"));
                                    experiencia.iCodCurriculumVitae = dr.GetInt32(dr.GetOrdinal("iCodCurriculumVitae"));
                                    experiencia.meses = dr.GetInt32(dr.GetOrdinal("meses"));
                                    experiencia.totalmeses = dr.GetInt32(dr.GetOrdinal("totalmeses"));
                                    lista.Add(experiencia);
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
        public Experiencia InsertarExperiencia(Experiencia entidad)
        {
            Experiencia Entidad = new Experiencia();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarExperiencia]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExperiencia", entidad.iCodExperiencia);
                        command.Parameters.AddWithValue("@vNombreEntidad", entidad.vNombreEntidad);
                        command.Parameters.AddWithValue("@vCargoServicio", entidad.vCargoServicio);
                        command.Parameters.AddWithValue("@vActividades", entidad.vActividades);
                        command.Parameters.AddWithValue("@vProductoServicio", entidad.vProductoServicio);                        
                        command.Parameters.AddWithValue("@dFechaInicio", entidad.dFechaInicio);
                        command.Parameters.AddWithValue("@dFechaFin", entidad.dFechaFin);
                        command.Parameters.AddWithValue("@vRutaArchivoConstancia", entidad.vRutaArchivoConstancia);
                        command.Parameters.AddWithValue("@iCodCurriculumVitae", entidad.iCodCurriculumVitae);
                        command.Parameters.AddWithValue("@iTipoExperiencia", entidad.iTipoExperiencia);
                        
                        command.Parameters.AddWithValue("@iOpcion", entidad.iOpcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())                                {

                                    Experiencia Experiencia = new Experiencia();

                                    Experiencia.iCodExperiencia = dr.GetInt32(dr.GetOrdinal("iCodExperiencia"));
                                    Entidad.iCodExperiencia = Experiencia.iCodExperiencia;
                                    Experiencia.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                    Entidad.vMensaje = Experiencia.vMensaje;
                                }
                            }
                        }
                    }
                    conection.Close();
                }
                return Entidad;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
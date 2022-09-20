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
    public class FormacionAcademicaDL
    {
        public FormacionAcademica InsertarFormacionAcademica(FormacionAcademica entidad)
        {
            FormacionAcademica Entidad = new FormacionAcademica();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_InsertarFormacionAcademica]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodFormacionAcademica", entidad.iCodFormacionAcademica);
                        command.Parameters.AddWithValue("@iCodNivel", entidad.iCodNivel);
                        command.Parameters.AddWithValue("@vCentroEstudios", entidad.vCentroEstudios);
                        command.Parameters.AddWithValue("@vEspecialidad", entidad.vEspecialidad);
                        command.Parameters.AddWithValue("@iCodExtensionista", entidad.iCodExtensionista);
                        command.Parameters.AddWithValue("@iCodCurriculumVitae", entidad.iCodCurriculumVitae);
                        command.Parameters.AddWithValue("@iOpcion", entidad.iOpcion);
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {

                                    FormacionAcademica FormacionAcademica = new FormacionAcademica();

                                    FormacionAcademica.iCodFormacionAcademica = dr.GetInt32(dr.GetOrdinal("iCodFormacionAcademica"));
                                    Entidad.iCodFormacionAcademica = FormacionAcademica.iCodFormacionAcademica;

                                    FormacionAcademica.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                    Entidad.vMensaje = FormacionAcademica.vMensaje;
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
        public List<FormacionAcademica> ListarFormacionAcademica(FormacionAcademica entidad)
        {
            List<FormacionAcademica> lista = new List<FormacionAcademica>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_ListarFormacionAcademica]", conection))
                    {

                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", entidad.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", entidad.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", entidad.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", entidad.pvSortOrder);

                        command.Parameters.AddWithValue("@iCodCurriculumVitae", entidad.iCodCurriculumVitae);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                FormacionAcademica FormacionAcademica;

                                while (dr.Read())
                                {

                                    FormacionAcademica = new FormacionAcademica();

                                    FormacionAcademica.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    FormacionAcademica.totalPaginas = dr.GetInt32(dr.GetOrdinal("iPageCount"));
                                    FormacionAcademica.paginaActual = dr.GetInt32(dr.GetOrdinal("iCurrentPage"));

                                    FormacionAcademica.iCodFormacionAcademica = dr.GetInt32(dr.GetOrdinal("iCodFormacionAcademica"));
                                    FormacionAcademica.iCodNivel = dr.GetInt32(dr.GetOrdinal("iCodNivel"));
                                    FormacionAcademica.descripcionnivel = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    FormacionAcademica.vCentroEstudios = dr.GetString(dr.GetOrdinal("vCentroEstudios"));
                                    FormacionAcademica.vEspecialidad = dr.GetString(dr.GetOrdinal("vEspecialidad"));                                    

                                    lista.Add(FormacionAcademica);
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

            return lista;
        }
    }
}
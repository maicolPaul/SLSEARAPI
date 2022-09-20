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


                   
                        command.Parameters.AddWithValue("@iCodNivel", entidad.iCodNivel);
                        command.Parameters.AddWithValue("@vCentroEstudios", entidad.vCentroEstudios);
                        command.Parameters.AddWithValue("@vEspecialidad", entidad.vEspecialidad);
                         


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
    }
}
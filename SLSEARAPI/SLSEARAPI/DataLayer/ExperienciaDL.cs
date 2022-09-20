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


                        command.Parameters.AddWithValue("@vNombreEntidad", entidad.vNombreEntidad);
                        command.Parameters.AddWithValue("@vCargoServicio", entidad.vCargoServicio);
                        command.Parameters.AddWithValue("@dFechaInicio", entidad.dFechaInicio);
                        command.Parameters.AddWithValue("@dFechaFin", entidad.dFechaFin);
                        command.Parameters.AddWithValue("@vRutaArchivoConstancia", entidad.vRutaArchivoConstancia);



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {

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
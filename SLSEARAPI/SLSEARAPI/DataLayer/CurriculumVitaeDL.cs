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
    public class CurriculumVitaeDL
    {
        public CurriculumVitae InsertarCurriculumVitae(CurriculumVitae entidad)
        {
            CurriculumVitae Entidad = new CurriculumVitae();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_InsertarCurriculumVitae]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;


                        command.Parameters.AddWithValue("@iCodExtensionista", entidad.iCodExtensionista);
                        command.Parameters.AddWithValue("@iCodFormacionAcademica", entidad.iCodFormacionAcademica);
                        command.Parameters.AddWithValue("@iCodExperiencia", entidad.iCodExperiencia);




                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {

                                    CurriculumVitae CurriculumVitae = new CurriculumVitae();

                                    CurriculumVitae.iCodCurriculumVitae = dr.GetInt32(dr.GetOrdinal("iCodCurriculumVitae"));
                                    Entidad.iCodCurriculumVitae = CurriculumVitae.iCodCurriculumVitae;

                                    CurriculumVitae.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                    Entidad.vMensaje = CurriculumVitae.vMensaje;

                      




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
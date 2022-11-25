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


                    using (SqlCommand command = new SqlCommand("[TB_Supervisio_CapCab]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodIdentificacion", entidad.iCodComponente);
                        command.Parameters.AddWithValue("@iCodFichaTecnica", entidad.iCodFichaTecnica);
                        command.Parameters.AddWithValue("@iCodComponente", entidad.iCodComponente);
                        command.Parameters.AddWithValue("@iCodActividad", entidad.iCodActividad);
                        command.Parameters.AddWithValue("@vObservaciongeneral", entidad.vObservaciongeneral);
                        command.Parameters.AddWithValue("@vRecomendacion", entidad.vRecomendacion);



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {

                                    entidad.iCodSuperCab = dr.GetInt32(dr.GetOrdinal("iCodHitos"));

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
    }
}
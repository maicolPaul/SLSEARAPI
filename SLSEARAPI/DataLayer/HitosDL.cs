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
    public class HitosDL
    {
        public Hito InsertarHito(Hito entidad)
        {       
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_InsertarHito]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                                         
                        command.Parameters.AddWithValue("@iCodComponente", entidad.iCodComponente);
                        command.Parameters.AddWithValue("@iCodActividad", entidad.iCodActividad);
                        command.Parameters.AddWithValue("@iCodHito", entidad.iCodHito);
                        command.Parameters.AddWithValue("@vTipo", entidad.vTipo);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {

                                    entidad.iCodHitos = dr.GetInt32(dr.GetOrdinal("iCodHitos"));
                                                              
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
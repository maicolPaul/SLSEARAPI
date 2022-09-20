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
    public class NivelDL
    {
        public List<Nivel> ListarNivel(Nivel entidad)
        {
            List<Nivel> Lista = new List<Nivel>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarNivel]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Lista = new List<Nivel>();
                                while (dr.Read())
                                {
                                    Nivel Nivel = new Nivel();
                                    Nivel.iCodNivel = dr.GetInt32(dr.GetOrdinal("iCodNivel"));
                                    Nivel.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));


                                    Lista.Add(Nivel);
                                }
                            }
                        }

                    }
                    conection.Close();
                }
                return Lista;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
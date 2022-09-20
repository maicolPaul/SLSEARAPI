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
    public class ProcedenciaDL
    {
        public List<Procedencia> ListarProcedencia(Procedencia entidad)
        {
            List<Procedencia> Lista = new List<Procedencia>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarProcedencia]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Lista = new List<Procedencia>();
                                while (dr.Read())
                                {
                                    Procedencia Procedencia = new Procedencia();
                                    Procedencia.iCodProcedencia = dr.GetInt32(dr.GetOrdinal("iCodProcedencia"));
                                    Procedencia.vProcedencia = dr.GetString(dr.GetOrdinal("vProcedencia"));


                                    Lista.Add(Procedencia);
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
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
    public class RequisitosDL
    {
        public List<Requisitos> ListarRequisitos(Requisitos entidad)
        {
            List<Requisitos> Lista = new List<Requisitos>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarRequisitos]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Lista = new List<Requisitos>();
                                while (dr.Read())
                                {
                                    Requisitos Requisitos = new Requisitos();
                                    Requisitos.iCodRequisito = dr.GetInt32(dr.GetOrdinal("iCodRequisito"));
                                    Requisitos.vRequisito = dr.GetString(dr.GetOrdinal("vRequisito"));


                                    Lista.Add(Requisitos);
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
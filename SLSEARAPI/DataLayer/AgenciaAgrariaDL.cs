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
    public class AgenciaAgrariaDL
    {
        public List<AgenciaAgraria> ListarAgenciaAgraria(AgenciaAgraria entidad)
        {
            List<AgenciaAgraria> Lista = new List<AgenciaAgraria>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarAgenciaAgraria]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodDirecGerencia", entidad.iCodDirecGerencia);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Lista = new List<AgenciaAgraria>();
                                while (dr.Read())
                                {
                                    AgenciaAgraria AgenciaAgraria = new AgenciaAgraria();
                                    AgenciaAgraria.iCodAgenciaAgraria = dr.GetInt32(dr.GetOrdinal("iCodAgenciaAgraria"));
                                    AgenciaAgraria.vAgencia = dr.GetString(dr.GetOrdinal("vAgencia"));


                                    Lista.Add(AgenciaAgraria);
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
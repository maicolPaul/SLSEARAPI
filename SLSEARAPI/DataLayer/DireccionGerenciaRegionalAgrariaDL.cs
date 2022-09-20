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
    public class DireccionGerenciaRegionalAgrariaDL
    {
        public List<DireccionGerenciaRegionalAgraria> ListarDireccionGerenciaRegionalAgraria(DireccionGerenciaRegionalAgraria entidad)
        {
            List<DireccionGerenciaRegionalAgraria> Lista = new List<DireccionGerenciaRegionalAgraria>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarDireccionGerenciaRegionalAgraria]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Lista = new List<DireccionGerenciaRegionalAgraria>();
                                while (dr.Read())
                                {
                                    DireccionGerenciaRegionalAgraria DireccionGerenciaRegionalAgraria = new DireccionGerenciaRegionalAgraria();
                                    DireccionGerenciaRegionalAgraria.iCodDirecGerencia = dr.GetInt32(dr.GetOrdinal("iCodDirecGerencia"));
                                    DireccionGerenciaRegionalAgraria.vNombre = dr.GetString(dr.GetOrdinal("vNombre"));


                                    Lista.Add(DireccionGerenciaRegionalAgraria);
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
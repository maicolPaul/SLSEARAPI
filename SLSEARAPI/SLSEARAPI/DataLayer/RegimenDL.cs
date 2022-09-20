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
    public class RegimenDL
    {
        public List<Regimen> ListarRegimen(Regimen entidad)
        {
            List<Regimen> Lista = new List<Regimen>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarRegimen]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Lista = new List<Regimen>();
                                while (dr.Read())
                                {
                                    Regimen Regimen = new Regimen();
                                    Regimen.iCodRegimen = dr.GetInt32(dr.GetOrdinal("iCodRegimen"));
                                    Regimen.vRegimen = dr.GetString(dr.GetOrdinal("vRegimen"));


                                    Lista.Add(Regimen);
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
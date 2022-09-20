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
    public class TipoEntidadDL
    {
        public List<TipoEntidad> ListarTipoEntidad(TipoEntidad entidad)
        {
            List<TipoEntidad> Lista = new List<TipoEntidad>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarTipoEntidad]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Lista = new List<TipoEntidad>();
                                while (dr.Read())
                                {
                                    TipoEntidad TipoEntidad = new TipoEntidad();
                                    TipoEntidad.iCodTipoEntidad = dr.GetInt32(dr.GetOrdinal("iCodTipoEntidad"));
                                    TipoEntidad.vTipoEntidad = dr.GetString(dr.GetOrdinal("vTipoEntidad"));


                                    Lista.Add(TipoEntidad);
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
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
    public class NivelDeInstruccionDL
    {
        public List<NivelDeInstruccion> ListarNivelDeInstruccion(NivelDeInstruccion entidad)
        {
            List<NivelDeInstruccion> Lista = new List<NivelDeInstruccion>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarNivelDeInstruccion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Lista = new List<NivelDeInstruccion>();
                                while (dr.Read())
                                {
                                    NivelDeInstruccion NivelDeInstruccion = new NivelDeInstruccion();
                                    NivelDeInstruccion.iCodNivelInstruccion = dr.GetInt32(dr.GetOrdinal("iCodNivelInstruccion"));
                                    NivelDeInstruccion.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));


                                    Lista.Add(NivelDeInstruccion);
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
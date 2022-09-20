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
    public class NombreArchivoDL
    {
        public List<NombreArchivo> ListarNombreArchivo(NombreArchivo entidad)
        {
            List<NombreArchivo> Lista = new List<NombreArchivo>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarNombreArchivo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Lista = new List<NombreArchivo>();
                                while (dr.Read())
                                {
                                    NombreArchivo NombreArchivo = new NombreArchivo();
                                    NombreArchivo.iCodNombreArchivo = dr.GetInt32(dr.GetOrdinal("iCodNombreArchivo"));
                                    NombreArchivo.vNombreArchivo = dr.GetString(dr.GetOrdinal("vNombreArchivo"));


                                    Lista.Add(NombreArchivo);
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
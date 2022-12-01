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
        public PorductorEjecucionTecnica InsertarProductorEje(PorductorEjecucionTecnica porductorEjecucionTecnica)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[SP_Insertar_ProductorEjec]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodComponente", porductorEjecucionTecnica.iCodComponente);
                        command.Parameters.AddWithValue("@iCodActividad", porductorEjecucionTecnica.iCodActividad);
                        command.Parameters.AddWithValue("@iCodProductor", porductorEjecucionTecnica.iCodProductor);
                        command.Parameters.AddWithValue("@dFechaCapa", porductorEjecucionTecnica.dFechaCapa);
                        command.Parameters.AddWithValue("@vTipo", porductorEjecucionTecnica.vTipo);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {

                                    porductorEjecucionTecnica.iCodProEje = dr.GetInt32(dr.GetOrdinal("iCodProEje"));

                                    porductorEjecucionTecnica.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                }
                            }
                        }

                    }
                    conection.Close();
                }
                return porductorEjecucionTecnica;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
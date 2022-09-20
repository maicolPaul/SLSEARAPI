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
    public class ListaChequeoRequisitosDL
    {
        public ListaChequeoRequisitos InsertarListaChequeoRequisitos(ListaChequeoRequisitos entidad)
        {
            ListaChequeoRequisitos Entidad = new ListaChequeoRequisitos();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_InsertarListaChequeoRequisitos]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;


                        command.Parameters.AddWithValue("@iCodExtensionista", entidad.iCodExtensionista );
                        command.Parameters.AddWithValue("@iCodRequisito", entidad.iCodRequisito);
                        command.Parameters.AddWithValue("@bCumple", entidad.bCumple);
               



                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {

                                    ListaChequeoRequisitos ListaChequeoRequisitos = new ListaChequeoRequisitos();

                                    ListaChequeoRequisitos.iCodListaChequeo = dr.GetInt32(dr.GetOrdinal("iCodListaChequeo"));
                                    Entidad.iCodListaChequeo = ListaChequeoRequisitos.iCodListaChequeo;

                                    ListaChequeoRequisitos.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                    Entidad.vMensaje = ListaChequeoRequisitos.vMensaje;






                                }

                            }
                        }

                    }
                    conection.Close();
                }
                return Entidad;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
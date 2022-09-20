using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Web;
using System.Web.Hosting;

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

        public ListaChequeoRequisitos ActualizarDardeBajaListaChequeoRequisitos(ListaChequeoRequisitos entidad)
        {
            ListaChequeoRequisitos Entidad = new ListaChequeoRequisitos();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ActualizarDardeBajaListaChequeoRequisitos]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodExtensionista", entidad.iCodExtensionista);
                        //command.Parameters.AddWithValue("@iCodRequisito", entidad.iCodRequisito);
                        //command.Parameters.AddWithValue("@bCumple", entidad.bCumple);

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

        public List<ListaChequeoRequisitos> ListarChequeoRequisitos(Extensionista extensionista)
        {
            List<ListaChequeoRequisitos> Lista = new List<ListaChequeoRequisitos>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarChequeoRequisitos]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodExtensionista", extensionista.iCodExtensionista);                        

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                ListaChequeoRequisitos listaChequeoRequisitos;
                                while (dr.Read())
                                {
                                    listaChequeoRequisitos = new ListaChequeoRequisitos();

                                    listaChequeoRequisitos.iCodListaChequeo = dr.GetInt32(dr.GetOrdinal("iCodListaChequeo"));
                                    listaChequeoRequisitos.iCodExtensionista = dr.GetInt32(dr.GetOrdinal("iCodExtensionista"));
                                    listaChequeoRequisitos.iCodRequisito = dr.GetInt32(dr.GetOrdinal("iCodRequisito"));
                                    listaChequeoRequisitos.bCumple = dr.GetBoolean(dr.GetOrdinal("bCumple"));
                                    
                                    Lista.Add(listaChequeoRequisitos);
                                                                        
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
        public bool GrabarRequisitos(List<ListaChequeoRequisitos> lista)
        {
            bool respuesta = false;
            try
            {
                foreach (ListaChequeoRequisitos item in lista)
                {
                    InsertarListaChequeoRequisitos(item);
                }
                respuesta = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return respuesta;
        }

   

    }
}
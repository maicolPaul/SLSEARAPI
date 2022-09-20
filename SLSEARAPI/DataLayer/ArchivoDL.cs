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
    public class ArchivoDL
    {
        public Archivo InsertarArchivo(Archivo entidad)
        {
            Archivo Entidad = new Archivo();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarArchivo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@icodExtensionista", entidad.icodExtensionista);
                        command.Parameters.AddWithValue("@iCodNombreArchivo", entidad.iCodNombreArchivo);
                        command.Parameters.AddWithValue("@vRutaArchivo", entidad.vRutaArchivo);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {

                                    Archivo Archivo = new Archivo();

                                    Archivo.iCodArchivos = dr.GetInt32(dr.GetOrdinal("iCodArchivos"));
                                    Entidad.iCodArchivos = Archivo.iCodArchivos;

                                    Archivo.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                    Entidad.vMensaje = Archivo.vMensaje;

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
        public List<Archivo> ListarArchivo(Archivo entidad)
        {
            List<Archivo> lista = new List<Archivo>();

            using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
            {
                conection.Open();


                using (SqlCommand command = new SqlCommand("[PA_ListarArchivo]", conection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@icodExtensionista", entidad.icodExtensionista);
                    command.Parameters.AddWithValue("@iCodNombreArchivo", entidad.iCodNombreArchivo);

                    using (SqlDataReader dr = command.ExecuteReader())
                    {
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                Archivo Archivo = new Archivo();

                                Archivo.iCodArchivos = dr.GetInt32(dr.GetOrdinal("iCodArchivos"));
                                
                                Archivo.vRutaArchivo = dr.GetString(dr.GetOrdinal("vRutaArchivo"));
                                lista.Add(Archivo);
                            }
                        }
                    }
                }
                conection.Close();
            }
                    return lista;
        }
        public Archivo EliminarArchivo(Archivo entidad)
        {
            Archivo Entidad = new Archivo();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EliminarArchivo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@icodExtensionista", entidad.icodExtensionista);
                        command.Parameters.AddWithValue("@iCodNombreArchivo", entidad.iCodNombreArchivo);
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {

                                    Archivo Archivo = new Archivo();

                                    Archivo.iCodArchivos = dr.GetInt32(dr.GetOrdinal("iCodArchivos"));
                                    Entidad.iCodArchivos = Archivo.iCodArchivos;

                                    Archivo.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                    Entidad.vMensaje = Archivo.vMensaje;

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
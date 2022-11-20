using Org.BouncyCastle.Cmp;
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
    public class ActaAlianzaEstrategicaDL
    {
        public List<Productor> ListarRepresentantes(Productor entidad)
        {
            List<Productor> lista = new List<Productor>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Listar_Representantes_Productor]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        
                        command.Parameters.AddWithValue("@iCodExtensionista", entidad.iCodExtensionista);
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Productor productor;
                                while (dr.Read())
                                {
                                    productor = new Productor();

                                    //productor.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    //productor.iPageCount = dr.GetInt32(dr.GetOrdinal("iPageCount"));
                                    //productor.piCurrentPage = dr.GetInt32(dr.GetOrdinal("iCurrentPage"));

                                    //productor.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    //productor.totalPaginas = dr.GetInt32(dr.GetOrdinal("iPageCount"));
                                    //productor.paginaActual = dr.GetInt32(dr.GetOrdinal("iCurrentPage"));

                                    productor.iCodProductor = dr.GetInt32(dr.GetOrdinal("iCodProductor"));

                                    productor.vApellidosNombres = dr.GetString(dr.GetOrdinal("vApellidosNombres"));

                                    productor.vDni = dr.GetString(dr.GetOrdinal("vDni"));
                                    productor.vCelular = dr.GetString(dr.GetOrdinal("vCelular"));
                                    productor.iEdad = dr.GetInt32(dr.GetOrdinal("iEdad"));
                                    productor.iSexo = dr.GetInt32(dr.GetOrdinal("iSexo"));
                                    productor.iPerteneceOrganizacion = dr.GetInt32(dr.GetOrdinal("iPerteneceOrganizacion"));

                                    productor.vNombreOrganizacion = dr.GetString(dr.GetOrdinal("vNombreOrganizacion"));
                                    productor.iEsRepresentante = dr.GetInt32(dr.GetOrdinal("iEsRepresentante"));

                                    productor.iCodExtensionista = dr.GetInt32(dr.GetOrdinal("iCodExtensionista"));
                                    lista.Add(productor);
                                }
                            }
                        }
                    }
                    conection.Close();
                }
                return lista;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<Productor> ListarProductor(Productor entidad)
        {
            List<Productor> lista = new List<Productor>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarProductor]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@piPageSize", entidad.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", entidad.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", entidad.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", entidad.pvSortOrder);                        

                        command.Parameters.AddWithValue("@iCodExtensionista", entidad.iCodExtensionista);
                        command.Parameters.AddWithValue("@iPerteneceOrganizacion", entidad.iPerteneceOrganizacion);
                        

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Productor productor;
                                while (dr.Read())
                                {
                                    productor = new Productor();

                                    //productor.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    //productor.iPageCount = dr.GetInt32(dr.GetOrdinal("iPageCount"));
                                    //productor.piCurrentPage = dr.GetInt32(dr.GetOrdinal("iCurrentPage"));

                                    productor.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    productor.totalPaginas = dr.GetInt32(dr.GetOrdinal("iPageCount"));
                                    productor.paginaActual = dr.GetInt32(dr.GetOrdinal("iCurrentPage"));

                                    productor.iCodProductor = dr.GetInt32(dr.GetOrdinal("iCodProductor"));

                                    productor.vApellidosNombres = dr.GetString(dr.GetOrdinal("vApellidosNombres"));

                                    productor.vDni=dr.GetString(dr.GetOrdinal("vDni"));
                                    productor.vCelular = dr.GetString(dr.GetOrdinal("vCelular"));
                                    productor.iEdad = dr.GetInt32(dr.GetOrdinal("iEdad"));
                                    productor.iSexo = dr.GetInt32(dr.GetOrdinal("iSexo"));
                                    productor.iPerteneceOrganizacion = dr.GetInt32(dr.GetOrdinal("iPerteneceOrganizacion"));

                                    productor.vNombreOrganizacion = dr.GetString(dr.GetOrdinal("vNombreOrganizacion"));
                                    productor.iEsRepresentante = dr.GetInt32(dr.GetOrdinal("iEsRepresentante"));

                                    productor.iCodExtensionista = dr.GetInt32(dr.GetOrdinal("iCodExtensionista"));
                                    productor.iRecibioCapacitacion = dr.GetInt32(dr.GetOrdinal("iRecibioCapacitacion"));
                                    productor.vNombreRepresentante = dr.GetString(dr.GetOrdinal("vNombreRepresentante"));
                                    productor.vRucOrg = dr.GetString(dr.GetOrdinal("vRucOrg"));
                                    productor.vTelefonoOrg = dr.GetString(dr.GetOrdinal("vTelefonoOrg"));
                                    productor.vCelularOrg = dr.GetString(dr.GetOrdinal("vCelularOrg"));
                                    productor.vDireccionOrg = dr.GetString(dr.GetOrdinal("vDireccionOrg"));
                                    productor.vCorreoElectronicoOrg = dr.GetString(dr.GetOrdinal("vCorreoElectronicoOrg"));
                                    productor.iCodTipoOrg = dr.GetInt32(dr.GetOrdinal("iCodTipoOrg")); 
                                    productor.cantidadmasculino= dr.GetInt32(dr.GetOrdinal("cantidadmasculino"));
                                    productor.cantidadfemenino = dr.GetInt32(dr.GetOrdinal("cantidadfemenino"));
                                    productor.promedio = dr.GetDecimal(dr.GetOrdinal("promedio"));
                                    productor.jovenes = dr.GetInt32(dr.GetOrdinal("jovenes"));
                                    productor.recibiocapacitacion = dr.GetInt32(dr.GetOrdinal("recibiocapacitacion"));
                                    productor.porfemenino = dr.GetDecimal(dr.GetOrdinal("porfemenino"));
                                    productor.porjovenes = dr.GetDecimal(dr.GetOrdinal("porjovenes"));
                                    productor.porrecibiocapacitacion = dr.GetDecimal(dr.GetOrdinal("porrecibiocapacitacion"));
                                    lista.Add(productor);
                                }
                            }
                        }
                    }
                    conection.Close();
                }
                return lista;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public Productor InsertarProductor(Productor entidad)
        {
            Productor Entidad = new Productor();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarProductor]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodProductor", entidad.iCodProductor);
                        command.Parameters.AddWithValue("@vApellidosNombres", entidad.vApellidosNombres);
                        command.Parameters.AddWithValue("@vDni", entidad.vDni);
                        command.Parameters.AddWithValue("@vCelular", entidad.vCelular);
                        command.Parameters.AddWithValue("@iEdad", entidad.iEdad);
                        command.Parameters.AddWithValue("@iSexo", entidad.iSexo);
                        command.Parameters.AddWithValue("@iPerteneceOrganizacion", entidad.iPerteneceOrganizacion);
                        command.Parameters.AddWithValue("@vNombreOrganizacion", entidad.vNombreOrganizacion);
                        command.Parameters.AddWithValue("@iRecibioCapacitacion", entidad.iRecibioCapacitacion);
                        command.Parameters.AddWithValue("@iEsRepresentante", entidad.iEsRepresentante);
                        command.Parameters.AddWithValue("@iCodExtensionista", entidad.iCodExtensionista);
                        command.Parameters.AddWithValue("@vNombreRepresentante", entidad.vNombreRepresentante);
                        command.Parameters.AddWithValue("@vRucOrg", entidad.vRucOrg);
                        command.Parameters.AddWithValue("@vTelefonoOrg", entidad.vTelefonoOrg);
                        command.Parameters.AddWithValue("@vCelularOrg", entidad.vCelularOrg);
                        command.Parameters.AddWithValue("@vDireccionOrg", entidad.vDireccionOrg);
                        command.Parameters.AddWithValue("@vCorreoElectronicoOrg", entidad.vCorreoElectronicoOrg);
                        command.Parameters.AddWithValue("@iCodTipoOrg", entidad.iCodTipoOrg);

                        command.Parameters.AddWithValue("@iOpcion", entidad.iOpcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                                                     

                                    entidad.iCodProductor = dr.GetInt32(dr.GetOrdinal("iCodProductor"));

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
    }
}
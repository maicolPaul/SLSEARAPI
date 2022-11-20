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
    public class FichaTecnicaDL
    {
        public List<FichaTecnica> ListarFichaTecnica(FichaTecnica fichaTecnicapar)
        {
            List<FichaTecnica> lista = new List<FichaTecnica>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Listar_FichaTecnica]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnicapar.iCodExtensionista);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                FichaTecnica fichatecnica;
                                while (dr.Read())
                                {
                                    fichatecnica = new FichaTecnica();
                                    fichatecnica.iCodFichaTecnica = dr.GetInt32(dr.GetOrdinal("iCodFichaTecnica"));
                                    fichatecnica.vNombreSearT1 = dr.GetString(dr.GetOrdinal("vNombreSearT1"));
                                    fichatecnica.vNaturalezaIntervencionT1= dr.GetString(dr.GetOrdinal("vNaturalezaIntervencionT1"));
                                    fichatecnica.vSubSectorT1 = dr.GetString(dr.GetOrdinal("vSubSectorT1"));
                                    fichatecnica.vCadenaProductivaT1 = dr.GetString(dr.GetOrdinal("vCadenaProductivaT1"));
                                    fichatecnica.vProcesoProductivaT1 = dr.GetString(dr.GetOrdinal("vProcesoProductivaT1"));
                                    fichatecnica.vLineaPrioritariaT1 = dr.GetString(dr.GetOrdinal("vLineaPrioritariaT1"));
                                    fichatecnica.vProductoServicioAmpliarT1 = dr.GetString(dr.GetOrdinal("vProductoServicioAmpliarT1"));
                                    fichatecnica.iCodUbigeoT1 = dr.GetString(dr.GetOrdinal("iCodUbigeoT1"));
                                    fichatecnica.vLocalidadT1 = dr.GetString(dr.GetOrdinal("vLocalidadT1"));
                                    fichatecnica.vZonaUTMT1 = dr.GetString(dr.GetOrdinal("vZonaUTMT1"));
                                    fichatecnica.vCoordenadasUTMNorteT1 = dr.GetString(dr.GetOrdinal("vCoordenadasUTMNorteT1"));
                                    fichatecnica.vCoordenadasUTMEsteT1 = dr.GetString(dr.GetOrdinal("vCoordenadasUTMEsteT1"));
                                    fichatecnica.dFechaInicioServicioT1 = dr.GetString(dr.GetOrdinal("dFechaInicioServicioT1"));
                                    fichatecnica.dFechaFinServicioT1 = dr.GetString(dr.GetOrdinal("dFechaFinServicioT1"));

                                    fichatecnica.vNombreEntidadProponenteT2 = dr.GetString(dr.GetOrdinal("vNombreEntidadProponenteT2"));
                                    fichatecnica.vNombreDireccionPerteneceT2 = dr.GetString(dr.GetOrdinal("vNombreDireccionPerteneceT2"));
                                    fichatecnica.vCorreoElectronicoT2 = dr.GetString(dr.GetOrdinal("vCorreoElectronicoT2"));
                                    fichatecnica.vNombreDirectorAgenciaAgrariaT2 = dr.GetString(dr.GetOrdinal("vNombreDirectorAgenciaAgrariaT2"));
                                    fichatecnica.vDireccionT2 = dr.GetString(dr.GetOrdinal("vDireccionT2"));
                                    fichatecnica.vTelefonoT2 = dr.GetString(dr.GetOrdinal("vTelefonoT2"));
                                    fichatecnica.vDireccionZonaAgroruralT2 = dr.GetString(dr.GetOrdinal("vDireccionZonaAgroruralT2"));

                                    lista.Add(fichatecnica);
                                }
                            }
                        }
                    }

                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }
        public List<TipoOrganizacion> ListarTipoOrganizacion()
        {
            List<TipoOrganizacion> lista = new List<TipoOrganizacion>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Listar_TipoOrganizacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                                                

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                TipoOrganizacion tipoProveedor;
                                while (dr.Read())
                                {
                                    tipoProveedor = new TipoOrganizacion();
                                    tipoProveedor.iCodTipoOrganizacion = dr.GetInt32(dr.GetOrdinal("iCodTipoOrganizacion"));
                                    tipoProveedor.vOrganizacion = dr.GetString(dr.GetOrdinal("vOrganizacion"));
                                    lista.Add(tipoProveedor);
                                }
                            }
                        }
                    }

                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }
        public List<LineaPrioritaria> ListarLineaPrioritaria(LineaPrioritaria lineaPrioritaria)
        {
            List<LineaPrioritaria> lista = new List<LineaPrioritaria>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarLineaPrioritaria]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodSector", lineaPrioritaria.iCodSector);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                LineaPrioritaria tipoProveedor;
                                while (dr.Read())
                                {
                                    tipoProveedor = new LineaPrioritaria();
                                    tipoProveedor.iCodLineaPriori = dr.GetInt32(dr.GetOrdinal("iCodLineaPriori"));
                                    tipoProveedor.vDescLineaPriori = dr.GetString(dr.GetOrdinal("vDescLineaPriori"));
                                    lista.Add(tipoProveedor);
                                }
                            }
                        }
                    }

                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }
        public List<CadenaProductivaAgraria> ListarCadenaProductivaAgraria()
        {
            List<CadenaProductivaAgraria> lista = new List<CadenaProductivaAgraria>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCadenaProductivaAgraria]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                CadenaProductivaAgraria tipoProveedor;
                                while (dr.Read())
                                {
                                    tipoProveedor = new CadenaProductivaAgraria();
                                    tipoProveedor.iCodProcesoCadPro = dr.GetInt32(dr.GetOrdinal("iCodProcesoCadPro"));
                                    tipoProveedor.vDescProcesoCadPro = dr.GetString(dr.GetOrdinal("vDescProcesoCadPro"));
                                    lista.Add(tipoProveedor);
                                }
                            }
                        }
                    }

                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }
        public List<Sector> ListarSector()
        {
            List<Sector> lista = new List<Sector>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarSector]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Sector tipoProveedor;
                                while (dr.Read())
                                {
                                    tipoProveedor = new Sector();
                                    tipoProveedor.iCodSector = dr.GetInt32(dr.GetOrdinal("iCodSector"));
                                    tipoProveedor.vDescSector = dr.GetString(dr.GetOrdinal("vDescSector"));
                                    lista.Add(tipoProveedor);
                                }
                            }
                        }
                    }

                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }

        public List<TipoProveedor> ListarTipoProveedor()
        {
            List<TipoProveedor> lista = new List<TipoProveedor>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarTipoProveedor]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                TipoProveedor tipoProveedor;
                                while (dr.Read())
                                {
                                    tipoProveedor = new TipoProveedor();
                                    tipoProveedor.iCodTipoProveedor = dr.GetInt32(dr.GetOrdinal("iCodTipoProveedor"));
                                    tipoProveedor.vProveedor = dr.GetString(dr.GetOrdinal("vProveedor"));
                                    lista.Add(tipoProveedor);
                                }
                            }
                        }
                    }
                   
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return lista;
        }
        public int RetornarDiferenciaMeses(FichaTecnica fichaTecnica)
        {
            int meses = 0;

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Calculardiferenciameses]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                                                
                        command.Parameters.AddWithValue("@dFechaInicioServicioT1", fichaTecnica.dFechaInicioServicioT1);
                        command.Parameters.AddWithValue("@dFechaFinServicioT1", fichaTecnica.dFechaFinServicioT1);                    

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    meses = dr.GetInt32(dr.GetOrdinal("meses"));                               
                                }
                            }
                        }
                    }
                    conection.Close();
                }
                return meses;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public FichaTecnica InsertarFichaTecnica(FichaTecnica fichaTecnica)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();
                                        
                    using (SqlCommand command = new SqlCommand("[PA_InsertarFichaTecnica]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodFichaTecnica", fichaTecnica.iCodFichaTecnica);
                        command.Parameters.AddWithValue("@vNombreSearT1", fichaTecnica.vNombreSearT1);
                        command.Parameters.AddWithValue("@vNaturalezaIntervencionT1", fichaTecnica.vNaturalezaIntervencionT1);
                        command.Parameters.AddWithValue("@vSubSectorT1", fichaTecnica.vSubSectorT1);
                        command.Parameters.AddWithValue("@vCadenaProductivaT1", fichaTecnica.vCadenaProductivaT1);
                        command.Parameters.AddWithValue("@vProcesoProductivaT1", fichaTecnica.vProcesoProductivaT1);
                        command.Parameters.AddWithValue("@vLineaPrioritariaT1", fichaTecnica.vLineaPrioritariaT1);
                        command.Parameters.AddWithValue("@vProductoServicioAmpliarT1", fichaTecnica.vProductoServicioAmpliarT1);
                        command.Parameters.AddWithValue("@iCodUbigeoT1", fichaTecnica.iCodUbigeoT1);
                        command.Parameters.AddWithValue("@vLocalidadT1", fichaTecnica.vLocalidadT1);
                        command.Parameters.AddWithValue("@vZonaUTMT1", fichaTecnica.vZonaUTMT1);
                        command.Parameters.AddWithValue("@vCoordenadasUTMNorteT1", fichaTecnica.vCoordenadasUTMNorteT1);
                        command.Parameters.AddWithValue("@vCoordenadasUTMEsteT1", fichaTecnica.vCoordenadasUTMEsteT1);
                        command.Parameters.AddWithValue("@dFechaInicioServicioT1", fichaTecnica.dFechaInicioServicioT1);
                        command.Parameters.AddWithValue("@dFechaFinServicioT1", fichaTecnica.dFechaFinServicioT1);
                        command.Parameters.AddWithValue("@iDuracionT1", fichaTecnica.iDuracionT1);

                        command.Parameters.AddWithValue("@vNombreEntidadProponenteT2", fichaTecnica.vNombreEntidadProponenteT2);
                        command.Parameters.AddWithValue("@vNombreDireccionPerteneceT2", fichaTecnica.vNombreDireccionPerteneceT2);
                        command.Parameters.AddWithValue("@vDireccionT2", fichaTecnica.vDireccionT2);
                        command.Parameters.AddWithValue("@vTelefonoT2", fichaTecnica.vTelefonoT2);
                        command.Parameters.AddWithValue("@vCorreoElectronicoT2", fichaTecnica.vCorreoElectronicoT2);
                        command.Parameters.AddWithValue("@vNombreDirectorAgenciaAgrariaT2", fichaTecnica.vNombreDirectorAgenciaAgrariaT2);
                        command.Parameters.AddWithValue("@vDireccionZonaAgroruralT2", fichaTecnica.vDireccionZonaAgroruralT2);

                        command.Parameters.AddWithValue("@iCodTipoPersoneriaT3", fichaTecnica.iCodTipoPersoneriaT3);
                        command.Parameters.AddWithValue("@vNombreRazonSocialProveedorT3", fichaTecnica.vNombreRazonSocialProveedorT3);
                        command.Parameters.AddWithValue("@vNombreRepresentanteLegalT3", fichaTecnica.vNombreRepresentanteLegalT3);
                        command.Parameters.AddWithValue("@vDniT3", fichaTecnica.vDniT3);
                        command.Parameters.AddWithValue("@vRucT3", fichaTecnica.vRucT3);
                        command.Parameters.AddWithValue("@vDireccionT3", fichaTecnica.vDireccionT3);
                        command.Parameters.AddWithValue("@vTelefonoT3", fichaTecnica.vTelefonoT3);
                        command.Parameters.AddWithValue("@vCelularT3", fichaTecnica.vCelularT3);
                        command.Parameters.AddWithValue("@vCorreoElectronicoT3", fichaTecnica.vCorreoElectronicoT3);
                        command.Parameters.AddWithValue("@vPaginaWebT3", fichaTecnica.vPaginaWebT3);
                        command.Parameters.AddWithValue("@vEpecialidadProveedorT3", fichaTecnica.vEpecialidadProveedorT3);
                        command.Parameters.AddWithValue("@iCodTipoProveedorT3", fichaTecnica.iCodTipoProveedorT3);
                        command.Parameters.AddWithValue("@iCodConvocatoria", fichaTecnica.iCodConvocatoria);
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);                        
                        command.Parameters.AddWithValue("@iOpcion", fichaTecnica.iOpcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    fichaTecnica.iCodFichaTecnica= dr.GetInt32(dr.GetOrdinal("iCodFichaTecnica"));                                    
                                    fichaTecnica.vMensaje = dr.GetString(dr.GetOrdinal("vmensaje"));                                    
                                }
                            }
                        }
                    }
                    conection.Close();
                }
                return fichaTecnica;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
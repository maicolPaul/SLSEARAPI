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
    public class CostoDL
    {
        public List<Actividad> ListarActividad(Actividad actividad)
        {
            List<Actividad> lista = new List<Actividad>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarActividadesPorExtensionista_Costo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", actividad.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", actividad.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", actividad.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", actividad.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodExtensionista", actividad.iCodExtensionista);
                        command.Parameters.AddWithValue("@iCodcomponente", actividad.iCodIdentificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    actividad = new Actividad();

                                    actividad.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    actividad.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    actividad.vActividad = dr.GetString(dr.GetOrdinal("vActividad"));
                                    actividad.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    actividad.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    actividad.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    actividad.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
                                    actividad.nTipoActividad = dr.GetInt32(dr.GetOrdinal("nTipoActividad"));
                                    actividad.bActivo = dr.GetBoolean(dr.GetOrdinal("bActivo"));
                                    actividad.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    lista.Add(actividad);
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

        public List<Costo> ListarCosto(Costo costo)
        {
            List<Costo> lista = new List<Costo>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", costo.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", costo.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", costo.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", costo.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodActividad", costo.iCodActividad);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    costo = new Costo();

                                    costo.iCodCosto = dr.GetInt32(dr.GetOrdinal("iCodCosto"));
                                    //costo.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    //costo.iCodComponente = dr.GetInt32(dr.GetOrdinal("iCodComponente"));
                                    costo.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    costo.iTipoMatServ = dr.GetInt32(dr.GetOrdinal("iTipoMatServ"));
                                    costo.TipoMatServ = dr.GetString(dr.GetOrdinal("TipoMatServ"));
                                    costo.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    costo.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    costo.iCantidad = dr.GetInt32(dr.GetOrdinal("iCantidad"));
                                    costo.dCostoUnitario = dr.GetDecimal(dr.GetOrdinal("dCostoUnitario"));
                                    costo.dFecha = dr.GetString(dr.GetOrdinal("dFecha"));
                                    costo.Estado = dr.GetString(dr.GetOrdinal("Estado"));
                                    costo.bActivo = dr.GetBoolean(dr.GetOrdinal("bActivo"));
                                    costo.iRecordCount = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    costo.iCodHito = dr.GetInt32(dr.GetOrdinal("iCodHito"));
                                    lista.Add(costo);
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

        public Costo InsertarCosto(Costo costo)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodCosto", costo.iCodCosto);
                        //command.Parameters.AddWithValue("@iCodIdentificacion", costo.iCodIdentificacion);
                        //command.Parameters.AddWithValue("@iCodComponente", costo.iCodComponente);
                        command.Parameters.AddWithValue("@iCodActividad", costo.iCodActividad);
                        command.Parameters.AddWithValue("@iTipoMatServ", costo.iTipoMatServ);
                        command.Parameters.AddWithValue("@vDescripcion", costo.vDescripcion);
                        command.Parameters.AddWithValue("@vUnidadMedida", costo.vUnidadMedida);
                        command.Parameters.AddWithValue("@iCantidad", costo.iCantidad);
                        command.Parameters.AddWithValue("@dCostoUnitario", costo.dCostoUnitario);
                        command.Parameters.AddWithValue("@dFecha", costo.dFecha);
                        command.Parameters.AddWithValue("@iopcion", costo.iopcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    costo.iCodCosto = dr.GetInt32(dr.GetOrdinal("iCodCosto"));
                                    costo.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return costo;
        }

        public DataTable ListarComponentesRpt(Cronograma cronograma)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarComponentesCabeceraCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", cronograma.iCodExtensionista);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        da.Fill(dataTable);

                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dataTable;
        }

        public DataTable ListarActividadesPorComponente(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarActividadesPorComponente2Costo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodComponente", actividad.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodExtensionista", actividad.iCodExtensionista);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        da.Fill(dataTable);
                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dataTable;
        }

        public DataTable ListarCostosPorActividad(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCostosPorActividad_RptCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodActividad", actividad.iCodActividad);
                        command.Parameters.AddWithValue("@iCodExtensionista", actividad.iCodExtensionista);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        da.Fill(dataTable);
                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dataTable;
        }

        public DataTable ListarCostosResumenPorComponente(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCostosResumenPorComponente_RptCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iTipoMatServ", actividad.iopcion);
                        command.Parameters.AddWithValue("@iCodComponente", actividad.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodExtensionista", actividad.iCodExtensionista);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        da.Fill(dataTable);
                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dataTable;
        }

        public DataTable ListarCostosResumenGeneral(Actividad actividad)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCostosResumenGeneral_RptCosto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iTipoMatServ", actividad.iopcion);
                        //command.Parameters.AddWithValue("@iCodComponente", actividad.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodExtensionista", actividad.iCodExtensionista);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        da.Fill(dataTable);
                    }
                    conection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dataTable;
        }

        public List<FichaTecnica> ListarfichaTecnicaRpt (FichaTecnica fichaTecnica)
        {
            List<FichaTecnica> lista = new List<FichaTecnica>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_FichaTecnicaRpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    fichaTecnica = new FichaTecnica();

                                    fichaTecnica.vNombreSearT1 = dr.GetString(dr.GetOrdinal("vNombreSearT1"));
                                    fichaTecnica.vNaturalezaIntervencionT1 = dr.GetString(dr.GetOrdinal("vNaturalezaIntervencionT1"));
                                    fichaTecnica.vSubSectorT1 = dr.GetString(dr.GetOrdinal("vSubSectorT1"));
                                    fichaTecnica.vCadenaProductivaT1 = dr.GetString(dr.GetOrdinal("vCadenaProductivaT1"));
                                    fichaTecnica.vProcesoProductivaT1 = dr.GetString(dr.GetOrdinal("vProcesoProductivaT1"));
                                    fichaTecnica.vLineaPrioritariaT1 = dr.GetString(dr.GetOrdinal("vLineaPrioritariaT1"));
                                    fichaTecnica.vProductoServicioAmpliarT1 = dr.GetString(dr.GetOrdinal("vProductoServicioAmpliarT1"));
                                    fichaTecnica.iCodUbigeoT1 = dr.GetString(dr.GetOrdinal("iCodUbigeoT1"));
                                    fichaTecnica.vNomDepartamento = dr.GetString(dr.GetOrdinal("vNomDepartamento"));
                                    fichaTecnica.vNomProvincia = dr.GetString(dr.GetOrdinal("vNomProvincia"));
                                    fichaTecnica.vNomDistrito = dr.GetString(dr.GetOrdinal("vNomDistrito"));
                                    fichaTecnica.vLocalidadT1 = dr.GetString(dr.GetOrdinal("vLocalidadT1"));
                                    fichaTecnica.vZonaUTMT1 = dr.GetString(dr.GetOrdinal("vZonaUTMT1"));
                                    fichaTecnica.vCoordenadasUTMNorteT1 = dr.GetString(dr.GetOrdinal("vCoordenadasUTMNorteT1"));
                                    fichaTecnica.vCoordenadasUTMEsteT1 = dr.GetString(dr.GetOrdinal("vCoordenadasUTMEsteT1"));
                                    fichaTecnica.dFechaInicioServicioT1 = dr.GetString(dr.GetOrdinal("dFechaInicioServicioT1"));
                                    fichaTecnica.dFechaFinServicioT1 = dr.GetString(dr.GetOrdinal("dFechaFinServicioT1"));
                                    fichaTecnica.TotalDias = dr.GetInt32(dr.GetOrdinal("TotalDias"));

                                    fichaTecnica.vNombreEntidadProponenteT2 = dr.GetString(dr.GetOrdinal("vNombreEntidadProponenteT2"));
                                    fichaTecnica.vCorreoElectronicoT2 = dr.GetString(dr.GetOrdinal("vCorreoElectronicoT2"));
                                    fichaTecnica.vNombreDireccionPerteneceT2 = dr.GetString(dr.GetOrdinal("vNombreDireccionPerteneceT2"));
                                    fichaTecnica.vNombreDirectorAgenciaAgrariaT2 = dr.GetString(dr.GetOrdinal("vNombreDirectorAgenciaAgrariaT2"));
                                    fichaTecnica.vDireccionT2 = dr.GetString(dr.GetOrdinal("vDireccionT2"));
                                    fichaTecnica.vDireccionZonaAgroruralT2 = dr.GetString(dr.GetOrdinal("vDireccionZonaAgroruralT2"));
                                    fichaTecnica.vTelefonoT2 = dr.GetString(dr.GetOrdinal("vTelefonoT2"));
                                    
                                    fichaTecnica.TipoPersoneriaT3 = dr.GetString(dr.GetOrdinal("TipoPersoneriaT3"));
                                    fichaTecnica.vTelefonoT3 = dr.GetString(dr.GetOrdinal("vTelefonoT3"));
                                    fichaTecnica.vCelularT3 = dr.GetString(dr.GetOrdinal("vCelularT3"));
                                    fichaTecnica.vNombreRazonSocialProveedorT3 = dr.GetString(dr.GetOrdinal("vNombreRazonSocialProveedorT3"));
                                    fichaTecnica.vCorreoElectronicoT3 = dr.GetString(dr.GetOrdinal("vCorreoElectronicoT3"));
                                    fichaTecnica.vNombreRepresentanteLegalT3 = dr.GetString(dr.GetOrdinal("vNombreRepresentanteLegalT3"));
                                    fichaTecnica.vPaginaWebT3 = dr.GetString(dr.GetOrdinal("vPaginaWebT3"));
                                    fichaTecnica.vRucT3 = dr.GetString(dr.GetOrdinal("vRucT3"));
                                    fichaTecnica.vEpecialidadProveedorT3 = dr.GetString(dr.GetOrdinal("vEpecialidadProveedorT3"));
                                    fichaTecnica.vDireccionT3 = dr.GetString(dr.GetOrdinal("vDireccionT3"));
                                    fichaTecnica.iCodTipoProveedorT3 = dr.GetInt32(dr.GetOrdinal("iCodTipoProveedorT3"));
                                    fichaTecnica.vProveedor = dr.GetString(dr.GetOrdinal("vProveedor"));

                                    lista.Add(fichaTecnica);
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

        public List<Productor> ListarProductorRpt(FichaTecnica fichaTecnica)
        {
            List<Productor> lista = new List<Productor>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_ProductoresRpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    Productor productor = new Productor();
                                    productor.Nro = dr.GetInt64(dr.GetOrdinal("Nro"));
                                    productor.vApellidosNombres = dr.GetString(dr.GetOrdinal("vApellidosNombres"));
                                    productor.vDni = dr.GetString(dr.GetOrdinal("vDni"));
                                    productor.vCelular = dr.GetString(dr.GetOrdinal("vCelular"));
                                    productor.iEdad = dr.GetInt32(dr.GetOrdinal("iEdad"));
                                    productor.vSexo = dr.GetString(dr.GetOrdinal("vSexo"));
                                    productor.vRecibioCapacitacion = dr.GetString(dr.GetOrdinal("vRecibioCapacitacion"));

                                    lista.Add(productor);
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

        public List<Productor> SP_Listar_OrganizacionesRpt(FichaTecnica fichaTecnica)
        {
            List<Productor> lista = new List<Productor>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_OrganizacionesRpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    Productor productor = new Productor();
                                    productor.vNombreOrganizacion = dr.GetString(dr.GetOrdinal("vNombreOrganizacion"));
                                    productor.vNombreRepresentante = dr.GetString(dr.GetOrdinal("vNombreRepresentante"));
                                    productor.vRucOrg = dr.GetString(dr.GetOrdinal("vRucOrg"));
                                    productor.vTelefonoOrg = dr.GetString(dr.GetOrdinal("vTelefonoOrg"));
                                    productor.vCelularOrg = dr.GetString(dr.GetOrdinal("vCelularOrg"));
                                    productor.vDireccionOrg = dr.GetString(dr.GetOrdinal("vDireccionOrg"));
                                    productor.vCorreoElectronicoOrg = dr.GetString(dr.GetOrdinal("vCorreoElectronicoOrg"));
                                    productor.iCodTipoOrg = dr.GetInt32(dr.GetOrdinal("iCodTipoOrg")); 
                                    productor.vOrganizacion = dr.GetString(dr.GetOrdinal("vOrganizacion"));
                                    lista.Add(productor);
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

        public List<Tecnologias> SP_Listar_TecnologiasRpt(FichaTecnica fichaTecnica)
        {
            List<Tecnologias> lista = new List<Tecnologias>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_TecnologiasRpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    Tecnologias tecnologia = new Tecnologias();
                                    tecnologia.vtecnologia1 = dr.GetString(dr.GetOrdinal("vtecnologia1"));
                                    tecnologia.vtecnologia2 = dr.GetString(dr.GetOrdinal("vtecnologia2"));
                                    tecnologia.vtecnologia3 = dr.GetString(dr.GetOrdinal("vtecnologia3"));
                                   
                                    lista.Add(tecnologia);
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
        
        public List<CausasIndirectas> SP_Listar_CausasDirectasIndirectasRpt(FichaTecnica fichaTecnica)
        {
            List<CausasIndirectas> lista = new List<CausasIndirectas>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_CausasDirectasIndirectasRpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    CausasIndirectas tecnologia = new CausasIndirectas();
                                    tecnologia.vDescrCausaInDirecta = dr.GetString(dr.GetOrdinal("vDescCausaIndirecta"));
                                    tecnologia.vDescrCausaDirecta = dr.GetString(dr.GetOrdinal("vDescrCausaDirecta"));
                                    tecnologia.vProblemaCentral = dr.GetString(dr.GetOrdinal("vProblemacentral"));
                                   
                                    lista.Add(tecnologia);
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

        public List<EfectoIndirecto> SP_Listar_EfectosDirectosIndirectosRpt(FichaTecnica fichaTecnica)
        {
            List<EfectoIndirecto> lista = new List<EfectoIndirecto>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[SP_Listar_EfectosDirectosIndirectosRpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    EfectoIndirecto tecnologia = new EfectoIndirecto();
                                    tecnologia.vDescEfectoIndirecto = dr.GetString(dr.GetOrdinal("vDescEfectoIndirecto"));
                                    tecnologia.vDescEfectoDirecto = dr.GetString(dr.GetOrdinal("vDescEfecto"));

                                    lista.Add(tecnologia);
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

        public List<Identificacion> PA_ListarDetallePobObj_Rpt(FichaTecnica fichaTecnica)
        {
            List<Identificacion> lista = new List<Identificacion>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarDetallePobObj_Rpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    Identificacion identificacion = new Identificacion();
                                    identificacion.vLimitaciones = dr.GetString(dr.GetOrdinal("vLimitaciones"));
                                    identificacion.vEstadoSituacional = dr.GetString(dr.GetOrdinal("vEstadoSituacional"));
                                    identificacion.vNumeroUnidadesProductivas = dr.GetString(dr.GetOrdinal("vNumeroUnidadesProductivas"));
                                    identificacion.vUnidadMedidaProductivas = dr.GetString(dr.GetOrdinal("vUnidadMedidaProductivas"));
                                    identificacion.vNumerosFamiliares = dr.GetInt32(dr.GetOrdinal("vNumerosFamiliares"));
                                    identificacion.vCantidad = dr.GetInt32(dr.GetOrdinal("vCantidad"));
                                    identificacion.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    identificacion.vRendimientoCadenaProductiva = dr.GetString(dr.GetOrdinal("vRendimientoCadenaProductiva"));
                                    identificacion.vGremios = dr.GetString(dr.GetOrdinal("vGremios"));
                                    identificacion.vObjetivoCentral = dr.GetString(dr.GetOrdinal("vObjetivoCentral"));
                                    lista.Add(identificacion);
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

        public List<Indicadores> PA_ListarIndicadores_Rpt(FichaTecnica fichaTecnica)
        {
            List<Indicadores> lista = new List<Indicadores>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarIndicadores_Rpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    Indicadores identificacion = new Indicadores();
                                    identificacion.TipoIndicador = dr.GetString(dr.GetOrdinal("TipoIndicador"));
                                    identificacion.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    identificacion.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    identificacion.vMedioVerificacion = dr.GetString(dr.GetOrdinal("vMedioVerificacion"));
                                    
                                    lista.Add(identificacion);
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

        public List<Componente> PA_ListarComponentes_Rpt(FichaTecnica fichaTecnica)
        {
            List<Componente> lista = new List<Componente>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarComponentes_Rpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);
                        command.Parameters.AddWithValue("@nTipoComponente", fichaTecnica.TotalDias);
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    Componente componente = new Componente();
                                    componente.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    componente.nTipoComponente = dr.GetInt32(dr.GetOrdinal("nTipoComponente"));
                                    componente.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    componente.vIndicador = dr.GetString(dr.GetOrdinal("vIndicador"));
                                    componente.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    componente.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    componente.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
                                    lista.Add(componente);
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

        public List<Componente> PA_ListarComponentes_Rpt1(FichaTecnica fichaTecnica)
        {
            List<Componente> lista = new List<Componente>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarComponentes_Rpt1]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", fichaTecnica.iCodExtensionista);
                        //command.Parameters.AddWithValue("@nTipoComponente", fichaTecnica.TotalDias);
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    Componente componente = new Componente();
                                    componente.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    componente.nTipoComponente = dr.GetInt32(dr.GetOrdinal("nTipoComponente"));
                                    componente.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    componente.vIndicador = dr.GetString(dr.GetOrdinal("vIndicador"));
                                    componente.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    componente.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    componente.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
                                    lista.Add(componente);
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

        public List<Actividad> ListarActividadesRpt(FichaTecnica fichaTecnica)
        {
            List<Actividad> lista = new List<Actividad>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarActividades_Rpt]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodComponenteDesc", fichaTecnica.iCodConvocatoria);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    Actividad actividad = new Actividad();

                                    actividad.vActividad = dr.GetString(dr.GetOrdinal("vActividad"));
                                    actividad.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    actividad.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    actividad.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    actividad.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
                                    lista.Add(actividad);
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

    }
}
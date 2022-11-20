using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace SLSEARAPI.DataLayer
{
    public class IdentificacionDL
    {        
        public List<Actividad> ListarActividades(Actividad actividad)
        {
            List<Actividad> lista = new List<Actividad>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarActividadesGeneral]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodExtensionista", actividad.iCodExtensionista);
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
                                    actividad.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    actividad.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
                                    actividad.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    actividad.nTipoActividad = dr.GetInt32(dr.GetOrdinal("nTipoActividad"));                                    

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
        public List<Componente> ListarComponentePorUsuario(Componente componente)
        {
            List<Componente> lista = new List<Componente>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarComponentesPorUsuario]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodIdentificacion", componente.iCodIdentificacion);
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {                                
                                while (dr.Read())
                                {
                                    componente = new Componente();
                                    componente.iCodComponente = dr.GetInt32(dr.GetOrdinal("iCodComponente"));
                                    componente.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));                                    
                                    componente.vIndicador = dr.GetString(dr.GetOrdinal("vIndicador"));

                                    componente.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    componente.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    componente.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
                                    componente.nTipoComponente = dr.GetInt32(dr.GetOrdinal("nTipoComponente"));
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
        public List<Indicadores> ListarIndicadores(Indicadores indicadores)
        {
            List<Indicadores> lista = new List<Indicadores>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarIndicadores]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodIdentificacion", indicadores.iCodIdentificacion);
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Indicadores tecnologia;
                                while (dr.Read())
                                {
                                    tecnologia = new Indicadores();
                                    tecnologia.iCodIndicador = dr.GetInt32(dr.GetOrdinal("iCodIndicador"));
                                    tecnologia.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    tecnologia.iCodigoIdentificador = dr.GetInt32(dr.GetOrdinal("iCodigoIdentificador"));
                                    tecnologia.vDescIdentificador = dr.GetString(dr.GetOrdinal("vDescIdentificador"));

                                    tecnologia.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    tecnologia.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    tecnologia.vMedioVerificacion = dr.GetString(dr.GetOrdinal("vMedioVerificacion"));

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

        public List<Tecnologias> ListarTecnologias(Tecnologias tecnologias)
        {
            List<Tecnologias> lista = new List<Tecnologias>();
            
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarTecnologias]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodIdentificacion", tecnologias.iCodIdentificacion);                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Tecnologias tecnologia;
                                while (dr.Read())
                                {
                                    tecnologia = new Tecnologias();
                                    tecnologia.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    tecnologia.vtecnologia1 = dr.GetString(dr.GetOrdinal("vTecnologia1"));
                                    tecnologia.vtecnologia2 = dr.GetString(dr.GetOrdinal("vTecnologia2"));
                                    tecnologia.vtecnologia3 = dr.GetString(dr.GetOrdinal("vTecnologia3"));                                    
                                    lista.Add(tecnologia);
                                }
                            }
                        }
                    }
                    conection.Close();
                }
                return lista;
            }
            catch (Exception)
            {
                throw;
            }
        }
        public List<Identificacion> ListarIdentificacion(Identificacion identificacion)
        {
            List<Identificacion> lista = new List<Identificacion>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarIdentificacion]", conection)){

                        command.CommandType = CommandType.StoredProcedure;                        
                        command.Parameters.AddWithValue("@iCodExtensionista", identificacion.iCodExtensionista);
                        //command.Parameters.AddWithValue("@tecnologiasxml", identificacion.tecnologiasxml);
                        //command.Parameters.AddWithValue("@iOpcion", identificacion.iOpcion);
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Identificacion identificaciondato;
                                while (dr.Read())
                                {
                                    identificaciondato = new Identificacion();
                                    identificacion.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    identificacion.vLimitaciones = dr.GetString(dr.GetOrdinal("vLimitaciones"));
                                    identificacion.vEstadoSituacional = dr.GetString(dr.GetOrdinal("vEstadoSituacional"));
                                    identificacion.vProblemacentral = dr.GetString(dr.GetOrdinal("vProblemacentral"));
                                    identificacion.vNumeroUnidadesProductivas = dr.GetString(dr.GetOrdinal("vNumeroUnidadesProductivas"));
                                    identificacion.vUnidadMedidaProductivas = dr.GetString(dr.GetOrdinal("vUnidadMedidaProductivas"));
                                    identificacion.vNumerosFamiliares = dr.GetInt32(dr.GetOrdinal("vNumerosFamiliares"));
                                    identificacion.vCantidad = dr.GetInt32(dr.GetOrdinal("vCantidad"));
                                    identificacion.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    identificacion.vRendimientoCadenaProductiva = dr.GetString(dr.GetOrdinal("vRendimientoCadenaProductiva"));
                                    identificacion.vGremios = dr.GetString(dr.GetOrdinal("vGremios"));
                                    identificacion.vObjetivoCentral = dr.GetString(dr.GetOrdinal("vObjetivoCentral"));
                                    identificacion.vDescComponente1 = dr.GetString(dr.GetOrdinal("vDescComponente1"));
                                    identificacion.vDescComponente2 = dr.GetString(dr.GetOrdinal("vDescComponente2"));
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
        public Identificacion InsertarIdentificacion(Identificacion identificacion)
        {
            //DataSet dataSet = new DataSet();
            //DataTable dttecnologias = new DataTable();
            //dttecnologias.Columns.Add("vtecnologias1");
            //dttecnologias.Columns.Add("vtecnologias2");
            //dttecnologias.Columns.Add("vtecnologias3");

            //foreach (Tecnologias item in identificacion.listatecnologias)
            //{
            //    DataRow dataRow = dttecnologias.NewRow();
            //    dataRow["vtecnologias1"] = item.vtecnologia1;
            //    dataRow["vtecnologias2"] = item.vtecnologia2;
            //    dataRow["vtecnologias3"] = item.vtecnologia3;
            //    dttecnologias.Rows.Add(dataRow);
            //}

            //dataSet.Tables.Add(dttecnologias);

            //identificacion.tecnologiasxml = dataSet.GetXml();


            //DataSet datasetindicadores = new DataSet();

            //DataTable dtindicadores = new DataTable();

            //dtindicadores.Columns.Add("iCodigoIdentificador");
            //dtindicadores.Columns.Add("vUnidadMedida");
            //dtindicadores.Columns.Add("vMeta");
            //dtindicadores.Columns.Add("vMedioVerificacion");

            //foreach (Indicadores item in identificacion.listaindicadores)
            //{
            //    DataRow dataRow = dtindicadores.NewRow();
            //    dataRow["iCodigoIdentificador"] = item.iCodigoIdentificador;
            //    dataRow["vUnidadMedida"] = item.vUnidadMedida;
            //    dataRow["vMeta"] = item.vMeta;
            //    dataRow["vMedioVerificacion"] = item.vMedioVerificacion;
            //    dtindicadores.Rows.Add(dataRow);
            //}

            //datasetindicadores.Tables.Add(dtindicadores);

            //identificacion.indicadoresxml = datasetindicadores.GetXml();

            //DataSet datasetcausasdirectas = new DataSet();

            //DataTable dtcausasdirectas = new DataTable();

            //dtcausasdirectas.Columns.Add("id");
            //dtcausasdirectas.Columns.Add("vdescrcausadirecta");

            //foreach (CausasDirectas item in identificacion.listacausasdirectas)
            //{
            //    DataRow dataRow = dtcausasdirectas.NewRow();

            //    dataRow["id"] = item.id;
            //    dataRow["vdescrcausadirecta"] = item.vdescrcausadirecta;

            //    dtcausasdirectas.Rows.Add(dataRow);
            //}

            //datasetcausasdirectas.Tables.Add(dtcausasdirectas);

            //identificacion.causasdirectasxml = datasetcausasdirectas.GetXml();

            //DataSet datasetcausasindirectas = new DataSet();

            //DataTable dtcausasindirectas = new DataTable();

            //dtcausasindirectas.Columns.Add("iCodCausaDirecta");
            //dtcausasindirectas.Columns.Add("vDescrCausaInDirecta");

            //foreach (CausasIndirectas item in identificacion.listacausasindirectas)
            //{
            //    DataRow dataRow = dtcausasindirectas.NewRow();

            //    dataRow["iCodCausaDirecta"] = item.iCodCausaDirecta;
            //    dataRow["vDescrCausaInDirecta"] = item.vDescrCausaInDirecta;

            //    dtcausasindirectas.Rows.Add(dataRow);
            //}

            //datasetcausasindirectas.Tables.Add(dtcausasindirectas);

            //identificacion.causasindirectaxml = datasetcausasindirectas.GetXml();

            //DataSet datasetefectosdirectos = new DataSet();

            //DataTable dtefectosdirectos = new DataTable();

            //dtefectosdirectos.Columns.Add("id");
            //dtefectosdirectos.Columns.Add("vdescefectodirecto");
            //if (identificacion.listaefectodirectos!=null)
            //{
            //    foreach (EfectoDirecto item in identificacion.listaefectodirectos)
            //    {
            //        DataRow dataRow = dtefectosdirectos.NewRow();

            //        dataRow["id"] = item.id;
            //        dataRow["vdescefectodirecto"] = item.vdescefectodirecto;

            //        dtefectosdirectos.Rows.Add(dataRow);
            //    }
            //}            

            //datasetefectosdirectos.Tables.Add(dtefectosdirectos);

            //identificacion.efectosdirectosxml = datasetefectosdirectos.GetXml();

            //DataSet datasetefectosindirectos = new DataSet();

            //DataTable dtefectosindirectos = new DataTable();

            //dtefectosindirectos.Columns.Add("iCodEfectoDirecto");
            //dtefectosindirectos.Columns.Add("vdescrefectoindirecto");

            //if (identificacion.listaefectoindirectos!=null)
            //{
            //    foreach (EfectoIndirecto item in identificacion.listaefectoindirectos)
            //    {
            //        DataRow dataRow = dtefectosindirectos.NewRow();

            //        dataRow["iCodEfectoDirecto"] = item.iCodEfectoDirecto;
            //        dataRow["vdescrefectoindirecto"] = item.vDescEfectoIndirecto;

            //        dtefectosindirectos.Rows.Add(dataRow);
            //    }
            //}            

            //datasetefectosindirectos.Tables.Add(dtefectosindirectos);

            //identificacion.efectosindirectosxml = datasetefectosindirectos.GetXml();

            //DataSet datasetcomponentes = new DataSet();

            //DataTable dtcomponentes = new DataTable();

            //dtcomponentes.Columns.Add("vIndicador");
            //dtcomponentes.Columns.Add("vUnidadMedida");
            //dtcomponentes.Columns.Add("vMeta");
            //dtcomponentes.Columns.Add("vMedio");
            //dtcomponentes.Columns.Add("nTipoComponente");

            //if (identificacion.listacomponente != null)
            //{
            //    foreach (Componente item in identificacion.listacomponente)
            //    {
            //        DataRow dataRow = dtcomponentes.NewRow();

            //        dataRow["vIndicador"] = item.vIndicador;
            //        dataRow["vUnidadMedida"] = item.vUnidadMedida;
            //        dataRow["vMeta"] = item.vMeta;
            //        dataRow["vMedio"] = item.vMedio;
            //        dataRow["nTipoComponente"] = item.nTipoComponente;

            //        dtcomponentes.Rows.Add(dataRow);
            //    }
            //}

            //datasetcomponentes.Tables.Add(dtcomponentes);

            //identificacion.componentesxml = datasetcomponentes.GetXml();

            //DataSet datasetactividades = new DataSet();

            //DataTable dtactividades = new DataTable();

            //dtactividades.Columns.Add("vActividad");
            //dtactividades.Columns.Add("vDescripcion");
            //dtactividades.Columns.Add("vUnidadMedida");
            //dtactividades.Columns.Add("vMeta");
            //dtactividades.Columns.Add("vMedio");
            //dtactividades.Columns.Add("nTipoActividad");

            //if (identificacion.listaactividad != null)
            //{
            //    foreach (Actividad item in identificacion.listaactividad)
            //    {
            //        DataRow dataRow = dtactividades.NewRow();

            //        dataRow["vActividad"] = item.vActividad;
            //        dataRow["vDescripcion"] = item.vDescripcion;
            //        dataRow["vUnidadMedida"] = item.vUnidadMedida;
            //        dataRow["vMeta"] = item.vMeta;
            //        dataRow["vMedio"] = item.vMedio;
            //        dataRow["nTipoActividad"] = item.nTipoActividad;
            //        dtactividades.Rows.Add(dataRow);
            //    }
            //}

            //datasetactividades.Tables.Add(dtactividades);

            //identificacion.actividadesxml = datasetactividades.GetXml();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Insertar_Identificacion]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodIdentificacion", identificacion.iCodIdentificacion);
                        command.Parameters.AddWithValue("@vLimitaciones", identificacion.vLimitaciones);
                        command.Parameters.AddWithValue("@vEstadoSituacional", identificacion.vEstadoSituacional);
                        command.Parameters.AddWithValue("@vProblemacentral", identificacion.vProblemacentral);
                        command.Parameters.AddWithValue("@vNumeroUnidadesProductivas", identificacion.vNumeroUnidadesProductivas);
                        command.Parameters.AddWithValue("@vUnidadMedidaProductivas", identificacion.vUnidadMedidaProductivas);
                        command.Parameters.AddWithValue("@vNumerosFamiliares", identificacion.vNumerosFamiliares);
                        command.Parameters.AddWithValue("@vCantidad", identificacion.vCantidad);
                        command.Parameters.AddWithValue("@vUnidadMedida", identificacion.vUnidadMedida);
                        command.Parameters.AddWithValue("@vRendimientoCadenaProductiva", identificacion.vRendimientoCadenaProductiva);
                        command.Parameters.AddWithValue("@vGremios", identificacion.vGremios);
                        command.Parameters.AddWithValue("@vObjetivoCentral", identificacion.vObjetivoCentral);
                        //command.Parameters.AddWithValue("@vdesccomponente1", identificacion.vDescComponente1);
                        //command.Parameters.AddWithValue("@vdesccomponente2", identificacion.vDescComponente2);

                        command.Parameters.AddWithValue("@iCodExtensionista", identificacion.iCodExtensionista);

                        //command.Parameters.AddWithValue("@tecnologiasxml", identificacion.tecnologiasxml);
                        //command.Parameters.AddWithValue("@indicadoresxml", identificacion.indicadoresxml);

                        //command.Parameters.AddWithValue("@causasdirectasxml", identificacion.causasdirectasxml);
                        //command.Parameters.AddWithValue("@causasindirectasxml", identificacion.causasindirectaxml);
                        //command.Parameters.AddWithValue("@componentesxml", identificacion.componentesxml);
                        //command.Parameters.AddWithValue("@actividadesxml", identificacion.actividadesxml);
                        command.Parameters.AddWithValue("@iOpcion", identificacion.iOpcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    identificacion.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    identificacion.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                }
                            }
                        }
                    }
                    conection.Close();
                }
                return identificacion;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<Componente> ListarComponentesPaginadoPorUsuario(Componente componente)
        {
            List<Componente> componentes = new List<Componente>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Listar_ComponentesPaginadoPorUsuario]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", componente.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", componente.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", componente.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", componente.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodIdentificacion", componente.iCodIdentificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    componente = new Componente();

                                    componente.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    componente.iCodComponente = dr.GetInt32(dr.GetOrdinal("iCodComponente"));
                                    componente.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    componente.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    componente.vIndicador = dr.GetString(dr.GetOrdinal("vIndicador"));
                                    componente.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    componente.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    componente.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
                                    componente.vMedio_ = dr.GetString(dr.GetOrdinal("vMedio_"));
                                    componente.nTipoComponente = dr.GetInt32(dr.GetOrdinal("nTipoComponente"));
                                    componente.vCorrelativo = dr.GetString(dr.GetOrdinal("CorrelativoComponente"));
                                    componente.vDescripcionCorta = dr.GetString(dr.GetOrdinal("vDescripcionCorta"));
                                    componentes.Add(componente);
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

            return componentes;
        }
        public List<Actividad> ListarActividadesPorComponente(Actividad actividad)
        {
            List<Actividad> lista = new List<Actividad>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Listar_Actividades_Por_Componente]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", actividad.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", actividad.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", actividad.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", actividad.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodIdentificacion", actividad.iCodIdentificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    actividad = new Actividad();

                                    actividad.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    actividad.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    actividad.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    actividad.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    actividad.vActividad = dr.GetString(dr.GetOrdinal("vActividad"));
                                    actividad.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    actividad.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    actividad.vMedio = dr.GetString(dr.GetOrdinal("vMedio"));
                                    actividad.vMedioCorta = dr.GetString(dr.GetOrdinal("vMedioCorta"));
                                    actividad.nTipoActividad = dr.GetInt32(dr.GetOrdinal("nTipoActividad"));
                                    actividad.vDescripcionCorta = dr.GetString(dr.GetOrdinal("vDescripcionCorta"));
                                    actividad.Correlativo = dr.GetInt32(dr.GetOrdinal("Correlativo"));
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
        public Actividad EliminarActividad(Actividad actividad)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Eliminar_Actividad]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodActividad", actividad.iCodActividad);                      

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    actividad = new Actividad();

                                    actividad.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    actividad.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));

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

            return actividad;
        }

        public Actividad InsertarActividad(Actividad actividad)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Insertar_Actividad]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodIdentificacion", actividad.iCodIdentificacion);
                        command.Parameters.AddWithValue("@vActividad", actividad.vActividad);
                        command.Parameters.AddWithValue("@vDescripcion", actividad.vDescripcion);
                        command.Parameters.AddWithValue("@vUnidadMedida", actividad.vUnidadMedida);
                        command.Parameters.AddWithValue("@vMeta", actividad.vMeta);
                        command.Parameters.AddWithValue("@vMedio", actividad.vMedio);
                        command.Parameters.AddWithValue("@nTipoActividad", actividad.nTipoActividad);                        

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    actividad = new Actividad();
                                    
                                    actividad.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    actividad.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));                                    
                                
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

            return actividad;
        }
        public Componente InsertarComponente(Componente componente)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Insertar_Componente]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodIdentificacion", componente.iCodIdentificacion);
                        command.Parameters.AddWithValue("@vDescripcion", componente.vDescripcion);
                        command.Parameters.AddWithValue("@vIndicador", componente.vIndicador);
                        command.Parameters.AddWithValue("@vUnidadMedida", componente.vUnidadMedida);
                        command.Parameters.AddWithValue("@vMeta", componente.vMeta);
                        command.Parameters.AddWithValue("@vMedio", componente.vMedio);
                        command.Parameters.AddWithValue("@nTipoComponente", componente.nTipoComponente);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    componente = new Componente();

                                    componente.iCodComponente = dr.GetInt32(dr.GetOrdinal("iCodComponente"));
                                    componente.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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
            return componente;
        }
        public Componente ActualizarComponente(Componente componente)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Editar_Componente]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodIdentificacion", componente.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodComponente", componente.iCodComponente);
                        command.Parameters.AddWithValue("@vDescripcion", componente.vDescripcion);
                        command.Parameters.AddWithValue("@vIndicador", componente.vIndicador);
                        command.Parameters.AddWithValue("@vUnidadMedida", componente.vUnidadMedida);
                        command.Parameters.AddWithValue("@vMeta", componente.vMeta);
                        command.Parameters.AddWithValue("@vMedio", componente.vMedio);                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    componente = new Componente();

                                    componente.iCodComponente = dr.GetInt32(dr.GetOrdinal("iCodComponente"));
                                    componente.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return componente;
        }
        public Componente EliminarComponente(Componente componente)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Eliminar_Componente]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;                        
                        command.Parameters.AddWithValue("@iCodComponente", componente.iCodComponente);              
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    componente = new Componente();

                                    componente.iCodComponente = dr.GetInt32(dr.GetOrdinal("iCodComponente"));
                                    componente.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return componente;
        }
        public Actividad ActualizarActividad(Actividad actividad)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Editar_Actividad]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodIdentificacion", actividad.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodActividad", actividad.iCodActividad);
                        command.Parameters.AddWithValue("@vDescripcion", actividad.vDescripcion);
                        command.Parameters.AddWithValue("@vUnidadMedida", actividad.vUnidadMedida);
                        command.Parameters.AddWithValue("@vMeta", actividad.vMeta);
                        command.Parameters.AddWithValue("@vMedio", actividad.vMedio);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    actividad = new Actividad();

                                    actividad.iCodActividad = dr.GetInt32(dr.GetOrdinal("iCodActividad"));
                                    actividad.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return actividad;
        }

        public Actividad ActividadCorrelativo(Actividad actividad)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Obtener_ActividadCorrelativo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodComponente", actividad.iCodIdentificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    actividad = new Actividad();

                                    actividad.vActividadCorrelativo = dr.GetString(dr.GetOrdinal("vActividadCorrelativo"));
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

            return actividad;
        }
        public Tecnologias InsertarTecnologia(Tecnologias tecnologias)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Insertar_Tecnologia]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodIdentificacion", tecnologias.iCodIdentificacion);
                        command.Parameters.AddWithValue("@vTecnologia1", tecnologias.vtecnologia1);
                        command.Parameters.AddWithValue("@vTecnologia2", tecnologias.vtecnologia2);
                        command.Parameters.AddWithValue("@vTecnologia3", tecnologias.vtecnologia3);                        

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    tecnologias = new Tecnologias();

                                    tecnologias.iCodTecnologia = dr.GetInt32(dr.GetOrdinal("iCodTecnologia"));
                                    tecnologias.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return tecnologias;
        }
        public Tecnologias EditarTecnologia(Tecnologias tecnologias)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Editar_Tecnologia]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodTecnologia", tecnologias.iCodTecnologia);
                        command.Parameters.AddWithValue("@iCodIdentificacion", tecnologias.iCodIdentificacion);
                        command.Parameters.AddWithValue("@vTecnologia1", tecnologias.vtecnologia1);
                        command.Parameters.AddWithValue("@vTecnologia2", tecnologias.vtecnologia2);
                        command.Parameters.AddWithValue("@vTecnologia3", tecnologias.vtecnologia3);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    tecnologias = new Tecnologias();

                                    tecnologias.iCodTecnologia = dr.GetInt32(dr.GetOrdinal("iCodTecnologia"));
                                    tecnologias.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return tecnologias;
        }
        public Tecnologias EliminarTecnologia(Tecnologias tecnologias)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Eliminar_Tecnologia]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodTecnologia", tecnologias.iCodTecnologia);              

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    tecnologias = new Tecnologias();

                                    tecnologias.iCodTecnologia = dr.GetInt32(dr.GetOrdinal("iCodTecnologia"));
                                    tecnologias.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return tecnologias;
        }

        public List<Tecnologias> ListarTecnologiasPaginado(Tecnologias tecnologias)
        {
            List<Tecnologias> lista = new List<Tecnologias>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Listar_Tecnologias_Paginado]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@piPageSize", tecnologias.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", tecnologias.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", tecnologias.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", tecnologias.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodIdentificacion", tecnologias.iCodIdentificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Tecnologias tecnologia;
                                while (dr.Read())
                                {
                                    tecnologia = new Tecnologias();
                                    tecnologia.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    tecnologia.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    tecnologia.iCodTecnologia = dr.GetInt32(dr.GetOrdinal("iCodTecnologia"));
                                    tecnologia.vtecnologia1 = dr.GetString(dr.GetOrdinal("vTecnologia1"));
                                    tecnologia.vtecnologia2 = dr.GetString(dr.GetOrdinal("vTecnologia2"));
                                    tecnologia.vtecnologia3 = dr.GetString(dr.GetOrdinal("vTecnologia3"));
                                    tecnologia.vtecnologia1Corta = dr.GetString(dr.GetOrdinal("vtecnologia1Corta"));
                                    tecnologia.vtecnologia2Corta = dr.GetString(dr.GetOrdinal("vtecnologia2Corta"));
                                    tecnologia.vtecnologia3Corta = dr.GetString(dr.GetOrdinal("vtecnologia3Corta"));
                                    lista.Add(tecnologia);
                                }
                            }
                        }
                    }
                    conection.Close();
                }
                return lista;
            }
            catch (Exception)
            {
                throw;
            }
        }
        public CausasDirectas InsertarCausasDirectas(CausasDirectas causasDirectas)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarCausasDirectas]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodIdentificacion", causasDirectas.iCodIdentificacion);
                        command.Parameters.AddWithValue("@vDescrCausaDirecta", causasDirectas.vdescrcausadirecta);                      

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    causasDirectas = new CausasDirectas();

                                    causasDirectas.iCodCausaDirecta = dr.GetInt32(dr.GetOrdinal("iCodCausaDirecta"));
                                    causasDirectas.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return causasDirectas;
        }
        public CausasDirectas EditarCausasDirectas(CausasDirectas causasDirectas)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EditarCausasDirectas]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodCausaDirecta", causasDirectas.iCodCausaDirecta);
                        command.Parameters.AddWithValue("@vDescrCausaDirecta", causasDirectas.vdescrcausadirecta);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    causasDirectas = new CausasDirectas();

                                    causasDirectas.iCodCausaDirecta = dr.GetInt32(dr.GetOrdinal("iCodCausaDirecta"));
                                    causasDirectas.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return causasDirectas;
        }
        public CausasDirectas EliminarCausasDirectas(CausasDirectas causasDirectas)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EliminarCausasDirectas]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodCausaDirecta", causasDirectas.iCodCausaDirecta);
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    causasDirectas = new CausasDirectas();

                                    causasDirectas.iCodCausaDirecta = dr.GetInt32(dr.GetOrdinal("iCodCausaDirecta"));
                                    causasDirectas.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return causasDirectas;
        }
        public List<CausasDirectas> ListarCausasDirectas(CausasDirectas causasDirectas)
        {
            List<CausasDirectas> lista = new List<CausasDirectas>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCausasDirectas]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@piPageSize", causasDirectas.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", causasDirectas.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", causasDirectas.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", causasDirectas.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodIdentificacion", causasDirectas.iCodIdentificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {                                
                                while (dr.Read())
                                {
                                    causasDirectas = new CausasDirectas();
                                    causasDirectas.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    causasDirectas.iCodCausaDirecta = dr.GetInt32(dr.GetOrdinal("iCodCausaDirecta"));
                                    causasDirectas.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    causasDirectas.vdescrcausadirecta = dr.GetString(dr.GetOrdinal("vDescrCausaDirecta"));
                                    causasDirectas.vDescrCausaDirectaCorta = dr.GetString(dr.GetOrdinal("vDescrCausaDirectaCorta"));
                                    lista.Add(causasDirectas);
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
        public CausasIndirectas InsertarCausasIndirectas(CausasIndirectas causasIndirectas)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarCausasIndirectas]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodIdentificacion", causasIndirectas.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodCausaDirecta", causasIndirectas.iCodCausaDirecta);
                        command.Parameters.AddWithValue("@vDescCausaIndirecta", causasIndirectas.vDescrCausaInDirecta);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    causasIndirectas = new CausasIndirectas();

                                    causasIndirectas.iCodCausaIndirecta = dr.GetInt32(dr.GetOrdinal("iCodCausaIndirecta"));
                                    causasIndirectas.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return causasIndirectas;
        }
        public CausasIndirectas EditarCausasIndirectas(CausasIndirectas causasIndirectas)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EditarCausasIndirectas]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        
                        command.Parameters.AddWithValue("@iCodCausaIndirecta", causasIndirectas.iCodCausaIndirecta);
                        command.Parameters.AddWithValue("@iCodIdentificacion", causasIndirectas.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodCausaDirecta", causasIndirectas.iCodCausaDirecta);
                        command.Parameters.AddWithValue("@vDescCausaIndirecta", causasIndirectas.vDescrCausaInDirecta);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    causasIndirectas = new CausasIndirectas();

                                    causasIndirectas.iCodCausaIndirecta = dr.GetInt32(dr.GetOrdinal("iCodCausaIndirecta"));
                                    causasIndirectas.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return causasIndirectas;
        }
        public CausasIndirectas EliminarCausasIndirectas(CausasIndirectas causasIndirectas)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EliminarCausasInDirectas]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodCausaIndirecta", causasIndirectas.iCodCausaIndirecta);
            
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    causasIndirectas = new CausasIndirectas();

                                    causasIndirectas.iCodCausaIndirecta = dr.GetInt32(dr.GetOrdinal("iCodCausaIndirecta"));
                                    causasIndirectas.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return causasIndirectas;
        }
        public List<CausasIndirectas> ListarCausasIndirectas(CausasIndirectas causasIndirectas)
        {
            List<CausasIndirectas> lista = new List<CausasIndirectas>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCausasIndirectas]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@piPageSize", causasIndirectas.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", causasIndirectas.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", causasIndirectas.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", causasIndirectas.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodIdentificacion", causasIndirectas.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodCausaDirecta", causasIndirectas.iCodCausaDirecta);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    causasIndirectas = new CausasIndirectas();
                                    causasIndirectas.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    causasIndirectas.iCodCausaIndirecta = dr.GetInt32(dr.GetOrdinal("iCodCausaIndirecta"));
                                    causasIndirectas.iCodCausaDirecta = dr.GetInt32(dr.GetOrdinal("iCodCausaDirecta"));
                                    causasIndirectas.iCodCausaIndirecta = dr.GetInt32(dr.GetOrdinal("iCodCausaIndirecta"));
                                    causasIndirectas.vDescrCausaInDirecta = dr.GetString(dr.GetOrdinal("vDescCausaIndirecta"));
                                    causasIndirectas.vDescCausaIndirectaCorta = dr.GetString(dr.GetOrdinal("vDescCausaIndirectaCorta"));

                                    lista.Add(causasIndirectas);
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
        public EfectoDirecto InsertarEfectoDirecto(EfectoDirecto efectoDirecto)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarEfectoDirecto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodIdentificacion", efectoDirecto.iCodIdentificacion);
                        command.Parameters.AddWithValue("@vDescEfecto", efectoDirecto.vDescEfecto);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    efectoDirecto = new EfectoDirecto();

                                    efectoDirecto.iCodEfecto = dr.GetInt32(dr.GetOrdinal("iCodEfecto"));
                                    efectoDirecto.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return efectoDirecto;
        }
        public EfectoDirecto EditarEfectoDirecto(EfectoDirecto efectoDirecto)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EditarEfectoDirecto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                                                
                        command.Parameters.AddWithValue("@iCodEfecto", efectoDirecto.iCodEfecto);
                        command.Parameters.AddWithValue("@iCodIdentificacion", efectoDirecto.iCodIdentificacion);
                        command.Parameters.AddWithValue("@vDescEfecto", efectoDirecto.vDescEfecto);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    efectoDirecto = new EfectoDirecto();

                                    efectoDirecto.iCodEfecto = dr.GetInt32(dr.GetOrdinal("iCodEfecto"));
                                    efectoDirecto.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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
            return efectoDirecto;
        }
        public EfectoDirecto EliminarEfectoDirecto(EfectoDirecto efectoDirecto)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EliminarEfectoDirecto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodEfecto", efectoDirecto.iCodEfecto);                   

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    efectoDirecto = new EfectoDirecto();

                                    efectoDirecto.iCodEfecto = dr.GetInt32(dr.GetOrdinal("iCodEfecto"));
                                    efectoDirecto.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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
            return efectoDirecto;
        }
        public List<EfectoDirecto> ListarEfectoDirecto(EfectoDirecto efectoDirecto)
        {
            List<EfectoDirecto> lista = new List<EfectoDirecto>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarEfectosDirectos]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@piPageSize", efectoDirecto.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", efectoDirecto.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", efectoDirecto.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", efectoDirecto.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodIdentificacion", efectoDirecto.iCodIdentificacion);                        

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    efectoDirecto = new EfectoDirecto();
                                    efectoDirecto.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    efectoDirecto.iCodEfecto = dr.GetInt32(dr.GetOrdinal("iCodEfecto"));
                                    efectoDirecto.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    efectoDirecto.vDescEfecto = dr.GetString(dr.GetOrdinal("vDescEfecto"));
                                    efectoDirecto.vDescEfectoCorta = dr.GetString(dr.GetOrdinal("vDescEfectoCorta"));

                                    lista.Add(efectoDirecto);
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
        public EfectoIndirecto InsertarEfectoIndirecto(EfectoIndirecto efectoIndirecto)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Insertar_EfectoIndirecto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodIdentificacion", efectoIndirecto.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodEfecto", efectoIndirecto.iCodEfectoDirecto);
                        command.Parameters.AddWithValue("@vDescEfectoIndirecto", efectoIndirecto.vDescEfectoIndirecto);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    efectoIndirecto = new EfectoIndirecto();

                                    efectoIndirecto.iCodEfectoIndirecto = dr.GetInt32(dr.GetOrdinal("iCodEfectoIndirecto"));
                                    efectoIndirecto.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return efectoIndirecto;
        }
        public EfectoIndirecto EditarEfectoIndirecto(EfectoIndirecto efectoIndirecto)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EditarEfectoIndirecto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodEfectoIndirecto", efectoIndirecto.iCodEfectoIndirecto);
                        command.Parameters.AddWithValue("@vDescEfectoIndirecto", efectoIndirecto.vDescEfectoIndirecto);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    efectoIndirecto = new EfectoIndirecto();

                                    efectoIndirecto.iCodEfectoIndirecto = dr.GetInt32(dr.GetOrdinal("iCodEfectoIndirecto"));
                                    efectoIndirecto.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return efectoIndirecto;
        }
        public EfectoIndirecto EliminarIndirecto(EfectoIndirecto efectoIndirecto)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EliminarEfectoIndirecto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodEfectoIndirecto", efectoIndirecto.iCodEfectoIndirecto);
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    efectoIndirecto = new EfectoIndirecto();

                                    efectoIndirecto.iCodEfectoIndirecto = dr.GetInt32(dr.GetOrdinal("iCodEfectoIndirecto"));
                                    efectoIndirecto.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return efectoIndirecto;
        }
        public List<EfectoIndirecto> ListarEfectoIndirecto(EfectoIndirecto efectoIndirecto)
        {
            List<EfectoIndirecto> lista = new List<EfectoIndirecto>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarEfectoIndirecto]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@piPageSize", efectoIndirecto.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", efectoIndirecto.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", efectoIndirecto.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", efectoIndirecto.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodIdentificacion", efectoIndirecto.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodEfecto", efectoIndirecto.iCodEfectoDirecto);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    efectoIndirecto = new EfectoIndirecto();
                                    efectoIndirecto.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    efectoIndirecto.iCodEfectoIndirecto = dr.GetInt32(dr.GetOrdinal("iCodEfectoIndirecto"));
                                    efectoIndirecto.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    efectoIndirecto.iCodEfectoDirecto = dr.GetInt32(dr.GetOrdinal("iCodEfecto"));
                                    efectoIndirecto.vDescEfectoIndirecto = dr.GetString(dr.GetOrdinal("vDescEfectoIndirecto"));
                                    efectoIndirecto.vDescEfectoIndirectoCorta = dr.GetString(dr.GetOrdinal("vDescEfectoIndirectoCorta"));

                                    lista.Add(efectoIndirecto);
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
        public Indicadores InsertarIndicador(Indicadores indicadores)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarIndicador]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodIdentificacion", indicadores.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodigoIdentificador", indicadores.iCodigoIdentificador);
                        command.Parameters.AddWithValue("@vUnidadMedida", indicadores.vUnidadMedida);
                        command.Parameters.AddWithValue("@vMeta", indicadores.vMeta);
                        command.Parameters.AddWithValue("@vMedioVerificacion", indicadores.vMedioVerificacion);


                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    indicadores = new Indicadores();

                                    indicadores.iCodIndicador = dr.GetInt32(dr.GetOrdinal("iCodIndicador"));
                                    indicadores.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return indicadores;
        }
        public Indicadores EditarIndicador(Indicadores indicadores)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EditarIndicador]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodIndicador", indicadores.iCodIndicador);
                        command.Parameters.AddWithValue("@iCodIdentificacion", indicadores.iCodIdentificacion);
                        command.Parameters.AddWithValue("@iCodigoIdentificador", indicadores.iCodigoIdentificador);
                        command.Parameters.AddWithValue("@vUnidadMedida", indicadores.vUnidadMedida);
                        command.Parameters.AddWithValue("@vMeta", indicadores.vMeta);
                        command.Parameters.AddWithValue("@vMedioVerificacion", indicadores.vMedioVerificacion);


                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    indicadores = new Indicadores();

                                    indicadores.iCodIndicador = dr.GetInt32(dr.GetOrdinal("iCodIndicador"));
                                    indicadores.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return indicadores;
        }

        public Indicadores EliminarIndicador(Indicadores indicadores)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_EliminarIndicador]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodIndicador", indicadores.iCodIndicador);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    indicadores = new Indicadores();

                                    indicadores.iCodIndicador = dr.GetInt32(dr.GetOrdinal("iCodIndicador"));
                                    indicadores.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return indicadores;
        }
        
        public Componente InsertarCompDescrip(Componente componente)
        {
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_InsertarCompDescrip]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodComponenteDesc", componente.iCodComponenteDesc);
                        command.Parameters.AddWithValue("@iCodIdentificacion", componente.iCodIdentificacion);
                        command.Parameters.AddWithValue("@vDescripcion", componente.vDescripcion);
                        command.Parameters.AddWithValue("@iOpcion", componente.iOpcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    componente = new Componente();

                                    componente.iCodComponenteDesc = dr.GetInt32(dr.GetOrdinal("iCodComponenteDesc"));
                                    componente.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
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

            return componente;
        }
        public List<Indicadores> ListarIndicadoresPaginado(Indicadores indicadores)
        {
            List<Indicadores> lista = new List<Indicadores>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarIndicador]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@piPageSize", indicadores.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", indicadores.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", indicadores.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", indicadores.pvSortOrder);
                        command.Parameters.AddWithValue("@iCodIdentificacion", indicadores.iCodIdentificacion);
                    
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    indicadores = new Indicadores();
                                    indicadores.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    indicadores.iCodIndicador = dr.GetInt32(dr.GetOrdinal("iCodIndicador"));
                                    indicadores.iCodIdentificacion = dr.GetInt32(dr.GetOrdinal("iCodIdentificacion"));
                                    indicadores.iCodigoIdentificador = dr.GetInt32(dr.GetOrdinal("iCodigoIdentificador"));
                                    indicadores.vUnidadMedida = dr.GetString(dr.GetOrdinal("vUnidadMedida"));
                                    indicadores.vMeta = dr.GetString(dr.GetOrdinal("vMeta"));
                                    indicadores.vMedioVerificacion = dr.GetString(dr.GetOrdinal("vMedioVerificacion"));
                                    indicadores.vDescIdentificador = dr.GetString(dr.GetOrdinal("vdescIdentificador"));
                                    lista.Add(indicadores);
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

        public List<Componente> ListarComponentesSelect(Componente component)
        {
            List<Componente> lista = new List<Componente>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarComponentesSelect]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodIdentificacion", component.iCodIdentificacion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    component = new Componente();
                                    component.iCodComponenteDesc = dr.GetInt32(dr.GetOrdinal("iCodComponenteDesc"));
                                    component.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    lista.Add(component);
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
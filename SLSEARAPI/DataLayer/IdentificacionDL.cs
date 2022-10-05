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

                    using (SqlCommand command = new SqlCommand("[PA_ListarIdentificacion]", conection))                    {

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
            DataSet dataSet = new DataSet();
            DataTable dttecnologias = new DataTable();
            dttecnologias.Columns.Add("vtecnologias1");
            dttecnologias.Columns.Add("vtecnologias2");
            dttecnologias.Columns.Add("vtecnologias3");

            foreach (Tecnologias item in identificacion.listatecnologias)
            {
                DataRow dataRow = dttecnologias.NewRow();
                dataRow["vtecnologias1"] = item.vtecnologia1;
                dataRow["vtecnologias2"] = item.vtecnologia2;
                dataRow["vtecnologias3"] = item.vtecnologia3;
                dttecnologias.Rows.Add(dataRow);
            }

            dataSet.Tables.Add(dttecnologias);

            identificacion.tecnologiasxml = dataSet.GetXml();


            DataSet datasetindicadores = new DataSet();

            DataTable dtindicadores = new DataTable();

            dtindicadores.Columns.Add("iCodigoIdentificador");
            dtindicadores.Columns.Add("vUnidadMedida");
            dtindicadores.Columns.Add("vMeta");
            dtindicadores.Columns.Add("vMedioVerificacion");

            foreach (Indicadores item in identificacion.listaindicadores)
            {
                DataRow dataRow = dtindicadores.NewRow();
                dataRow["iCodigoIdentificador"] = item.iCodigoIdentificador;
                dataRow["vUnidadMedida"] = item.vUnidadMedida;
                dataRow["vMeta"] = item.vMeta;
                dataRow["vMedioVerificacion"] = item.vMedioVerificacion;
                dtindicadores.Rows.Add(dataRow);
            }

            datasetindicadores.Tables.Add(dtindicadores);

            identificacion.indicadoresxml = datasetindicadores.GetXml();

            DataSet datasetcausasdirectas = new DataSet();

            DataTable dtcausasdirectas = new DataTable();

            dtcausasdirectas.Columns.Add("id");
            dtcausasdirectas.Columns.Add("vdescrcausadirecta");

            //foreach (CausasDirectas item in identificacion.listacausasdirectas)
            //{
            //    DataRow dataRow = dtcausasdirectas.NewRow();

            //    dataRow["id"] = item.id;
            //    dataRow["vdescrcausadirecta"] = item.vdescrcausadirecta;

            //    dtcausasdirectas.Rows.Add(dataRow);
            //}

            datasetcausasdirectas.Tables.Add(dtcausasdirectas);

            identificacion.causasdirectasxml = datasetcausasdirectas.GetXml();

            DataSet datasetcausasindirectas = new DataSet();

            DataTable dtcausasindirectas = new DataTable();

            dtcausasindirectas.Columns.Add("iCodCausaDirecta");
            dtcausasindirectas.Columns.Add("vDescrCausaInDirecta");

            //foreach (CausasIndirectas item in identificacion.listacausasindirectas)
            //{
            //    DataRow dataRow = dtcausasindirectas.NewRow();

            //    dataRow["iCodCausaDirecta"] = item.iCodCausaDirecta;
            //    dataRow["vDescrCausaInDirecta"] = item.vDescrCausaInDirecta;

            //    dtcausasindirectas.Rows.Add(dataRow);
            //}

            datasetcausasindirectas.Tables.Add(dtcausasindirectas);

            identificacion.causasindirectaxml = datasetcausasindirectas.GetXml();

            DataSet datasetefectosdirectos = new DataSet();

            DataTable dtefectosdirectos = new DataTable();

            dtefectosdirectos.Columns.Add("id");
            dtefectosdirectos.Columns.Add("vdescefectodirecto");
            if (identificacion.listaefectodirectos!=null)
            {
                foreach (EfectoDirecto item in identificacion.listaefectodirectos)
                {
                    DataRow dataRow = dtefectosdirectos.NewRow();

                    dataRow["id"] = item.id;
                    dataRow["vdescefectodirecto"] = item.vdescefectodirecto;

                    dtefectosdirectos.Rows.Add(dataRow);
                }
            }            

            datasetefectosdirectos.Tables.Add(dtefectosdirectos);

            identificacion.efectosdirectosxml = datasetefectosdirectos.GetXml();

            DataSet datasetefectosindirectos = new DataSet();

            DataTable dtefectosindirectos = new DataTable();

            dtefectosindirectos.Columns.Add("iCodEfectoDirecto");
            dtefectosindirectos.Columns.Add("vdescrefectoindirecto");

            if (identificacion.listaefectoindirectos!=null)
            {
                foreach (EfectoIndirecto item in identificacion.listaefectoindirectos)
                {
                    DataRow dataRow = dtefectosindirectos.NewRow();

                    dataRow["iCodEfectoDirecto"] = item.iCodEfectoDirecto;
                    dataRow["vdescrefectoindirecto"] = item.vDescEfectoIndirecto;

                    dtefectosindirectos.Rows.Add(dataRow);
                }
            }            

            datasetefectosindirectos.Tables.Add(dtefectosindirectos);

            identificacion.efectosindirectosxml = datasetefectosindirectos.GetXml();

            DataSet datasetcomponentes = new DataSet();

            DataTable dtcomponentes = new DataTable();

            dtcomponentes.Columns.Add("vIndicador");
            dtcomponentes.Columns.Add("vUnidadMedida");
            dtcomponentes.Columns.Add("vMeta");
            dtcomponentes.Columns.Add("vMedio");
            dtcomponentes.Columns.Add("nTipoComponente");

            if (identificacion.listacomponente != null)
            {
                foreach (Componente item in identificacion.listacomponente)
                {
                    DataRow dataRow = dtcomponentes.NewRow();

                    dataRow["vIndicador"] = item.vIndicador;
                    dataRow["vUnidadMedida"] = item.vUnidadMedida;
                    dataRow["vMeta"] = item.vMeta;
                    dataRow["vMedio"] = item.vMedio;
                    dataRow["nTipoComponente"] = item.nTipoComponente;

                    dtcomponentes.Rows.Add(dataRow);
                }
            }

            datasetcomponentes.Tables.Add(dtcomponentes);

            identificacion.componentesxml = datasetcomponentes.GetXml();

            DataSet datasetactividades = new DataSet();

            DataTable dtactividades = new DataTable();

            dtactividades.Columns.Add("vActividad");
            dtactividades.Columns.Add("vDescripcion");
            dtactividades.Columns.Add("vUnidadMedida");
            dtactividades.Columns.Add("vMeta");
            dtactividades.Columns.Add("vMedio");
            dtactividades.Columns.Add("nTipoActividad");

            if (identificacion.listaactividad != null)
            {
                foreach (Actividad item in identificacion.listaactividad)
                {
                    DataRow dataRow = dtactividades.NewRow();

                    dataRow["vActividad"] = item.vActividad;
                    dataRow["vDescripcion"] = item.vDescripcion;
                    dataRow["vUnidadMedida"] = item.vUnidadMedida;
                    dataRow["vMeta"] = item.vMeta;
                    dataRow["vMedio"] = item.vMedio;
                    dataRow["nTipoActividad"] = item.nTipoActividad;
                    dtactividades.Rows.Add(dataRow);
                }
            }

            datasetactividades.Tables.Add(dtactividades);

            identificacion.actividadesxml = datasetactividades.GetXml();

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
                        command.Parameters.AddWithValue("@vdesccomponente1", identificacion.vDescComponente1);
                        command.Parameters.AddWithValue("@vdesccomponente2", identificacion.vDescComponente2);

                        command.Parameters.AddWithValue("@iCodExtensionista", identificacion.iCodExtensionista);

                        command.Parameters.AddWithValue("@tecnologiasxml", identificacion.tecnologiasxml);
                        command.Parameters.AddWithValue("@indicadoresxml", identificacion.indicadoresxml);

                        command.Parameters.AddWithValue("@causasdirectasxml", identificacion.causasdirectasxml);
                        command.Parameters.AddWithValue("@causasindirectasxml", identificacion.causasindirectaxml);
                        command.Parameters.AddWithValue("@componentesxml", identificacion.componentesxml);
                        command.Parameters.AddWithValue("@actividadesxml", identificacion.actividadesxml);

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
    }
}
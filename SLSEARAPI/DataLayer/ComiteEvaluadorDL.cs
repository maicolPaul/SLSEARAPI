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
    public class ComiteEvaluadorDL
    {
        public ComiteEvaluador InsertarComiteEvaluador(ComiteEvaluador comiteEvaluador)
        {
			try
			{
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_Insertar_ComiteEvaluador]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@iCodComiteEvaluador", comiteEvaluador.iCodComiteEvaluador);                        
                        command.Parameters.AddWithValue("@vNombres", comiteEvaluador.vNombres);
                        command.Parameters.AddWithValue("@vApellidoPat", comiteEvaluador.vApellidoPat);
                        command.Parameters.AddWithValue("@vApellidoMat", comiteEvaluador.vApellidoMat);
                        command.Parameters.AddWithValue("@iCodTipoDoc", comiteEvaluador.iCodTipoDoc);
                        command.Parameters.AddWithValue("@vNroDocumento", comiteEvaluador.vNroDocumento);
                        command.Parameters.AddWithValue("@vCodUbigeo", comiteEvaluador.vCodUbigeo);
                        command.Parameters.AddWithValue("@iCodCargo", comiteEvaluador.iCodCargo);
                        command.Parameters.AddWithValue("@vCelular", comiteEvaluador.vCelular);
                        command.Parameters.AddWithValue("@vCorreo", comiteEvaluador.vCorreo);
                        command.Parameters.AddWithValue("@iCodArchivos", comiteEvaluador.iCodArchivos);
                        command.Parameters.AddWithValue("@iopcion", comiteEvaluador.iopcion);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {

                                if (dr.Read())
                                {
                                    comiteEvaluador.iCodComiteEvaluador = dr.GetInt32(dr.GetOrdinal("iCodComiteEvaluador"));
                                    comiteEvaluador.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                }
                            }
                        }
                    }
                    conection.Close();
                }
            }
			catch (Exception)
			{

				throw;
			}

			return comiteEvaluador;
        }
        public List<ComiteEvaluador> ListarComiteEvaluador(ComiteEvaluador comiteEvaluador)
        {
            List<ComiteEvaluador> lista = new List<ComiteEvaluador>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarComiteEvaluador]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@piPageSize", comiteEvaluador.piPageSize);
                        command.Parameters.AddWithValue("@piCurrentPage", comiteEvaluador.piCurrentPage);
                        command.Parameters.AddWithValue("@pvSortColumn", comiteEvaluador.pvSortColumn);
                        command.Parameters.AddWithValue("@pvSortOrder", comiteEvaluador.pvSortOrder);                      

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    comiteEvaluador = new ComiteEvaluador();
                                    comiteEvaluador.totalRegistros = dr.GetInt32(dr.GetOrdinal("iRecordCount"));
                                    comiteEvaluador.iCodComiteEvaluador = dr.GetInt32(dr.GetOrdinal("iCodComiteEvaluador"));
                                    comiteEvaluador.vNombres = dr.GetString(dr.GetOrdinal("vNombres"));
                                    comiteEvaluador.vApellidoPat = dr.GetString(dr.GetOrdinal("vApellidoPat"));
                                    comiteEvaluador.vApellidoMat = dr.GetString(dr.GetOrdinal("vApellidoMat"));
                                    comiteEvaluador.iCodTipoDoc = dr.GetInt32(dr.GetOrdinal("iCodTipoDoc"));
                                    comiteEvaluador.vNroDocumento = dr.GetString(dr.GetOrdinal("vNroDocumento"));
                                    comiteEvaluador.vCodUbigeo = dr.GetString(dr.GetOrdinal("vCodUbigeo"));
                                    comiteEvaluador.vNomDistrito = dr.GetString(dr.GetOrdinal("vNomDistrito"));
                                    comiteEvaluador.vCodProvincia = dr.GetString(dr.GetOrdinal("vCodProvincia"));
                                    comiteEvaluador.vNomProvincia = dr.GetString(dr.GetOrdinal("vNomProvincia"));
                                    comiteEvaluador.vCodDepartamento = dr.GetString(dr.GetOrdinal("vCodDepartamento"));
                                    comiteEvaluador.vNomDepartamento = dr.GetString(dr.GetOrdinal("vNomDepartamento"));                                    
                                    comiteEvaluador.iCodCargo = dr.GetInt32(dr.GetOrdinal("iCodCargo"));
                                    comiteEvaluador.vDescripcionCargo = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                    comiteEvaluador.vCelular = dr.GetString(dr.GetOrdinal("vCelular"));
                                    comiteEvaluador.vCorreo = dr.GetString(dr.GetOrdinal("vCorreo"));
                                    comiteEvaluador.iCodArchivos = dr.GetInt32(dr.GetOrdinal("iCodArchivos"));
                                    comiteEvaluador.Estado = dr.GetString(dr.GetOrdinal("Estado"));
                                    lista.Add(comiteEvaluador);
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
        public List<Cargo> ListarCargo()
        {
            List<Cargo> lista = new List<Cargo>();

            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_ListarCargo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                  
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Cargo cargo;
                                while (dr.Read())
                                {
                                    cargo = new Cargo();
                                    cargo.iCodCargo = dr.GetInt32(dr.GetOrdinal("iCodCargo"));
                                    cargo.vDescripcion = dr.GetString(dr.GetOrdinal("vDescripcion"));
                                  
                                    lista.Add(cargo);
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
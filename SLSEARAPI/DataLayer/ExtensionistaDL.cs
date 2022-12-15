using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace SLSEARAPI.DataLayer
{
    public class ExtensionistaDL
    {
        public static string GetSHA256(string str)
        {
            SHA256 sha256 = SHA256Managed.Create();
            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] stream = null;
            StringBuilder sb = new StringBuilder();
            stream = sha256.ComputeHash(encoding.GetBytes(str));
            for (int i = 0; i < stream.Length; i++) sb.AppendFormat("{0:x2}", stream[i]);
            return sb.ToString();
        }
        public Extensionista ActualizarPropuestaExtensionista(Extensionista entidad)
        {
            Extensionista Entidad = new Extensionista();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();
                                        
                    using (SqlCommand command = new SqlCommand("[PA_ActualizarPropuestaExtensionista]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        
                        command.Parameters.AddWithValue("@iCodExtensionista", entidad.iCodExtensionista);                        
                        command.Parameters.AddWithValue("@vNombrePropuesta", entidad.vNombrePropuesta);

                        command.CommandTimeout = 0;
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {

                                    Extensionista Extensionista = new Extensionista();

                                    Extensionista.iCodExtensionista = dr.GetInt32(dr.GetOrdinal("iCodExtensionista"));
                                    Entidad.iCodExtensionista = Extensionista.iCodExtensionista;

                                    Extensionista.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                    Entidad.vMensaje = Extensionista.vMensaje;
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

        public Extensionista InsertarExtensionista(Extensionista entidad)
        {
            Extensionista Entidad = new Extensionista();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    entidad.vClave = GetSHA256(entidad.vClave);
                    using (SqlCommand command = new SqlCommand("[PA_InsertarExtensionista]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                         command.Parameters.AddWithValue("@vRazonSocial",entidad.vRazonSocial);
                         command.Parameters.AddWithValue("@vInicialesSiglas",entidad.vInicialesSiglas);
                         command.Parameters.AddWithValue("@iCodProcedencia",entidad.iCodProcedencia);
                         command.Parameters.AddWithValue("@iCodRegimen",entidad.iCodRegimen);
                         command.Parameters.AddWithValue("@iCodTipoEntidad",entidad.iCodTipoEntidad);
                         command.Parameters.AddWithValue("@vOtros",entidad.vOtros);
                         command.Parameters.AddWithValue("@vRucEmpresa",entidad.vRucEmpresa);
                         command.Parameters.AddWithValue("@vDomicilioFiscal",entidad.vDomicilioFiscal);
                         command.Parameters.AddWithValue("@vCodDepartamento",entidad.vCodDepartamento);
                         command.Parameters.AddWithValue("@vCodProvincia",entidad.vCodProvincia);
                         command.Parameters.AddWithValue("@vCodDistrito",entidad.vCodDistrito);
                         command.Parameters.AddWithValue("@vTelefono",entidad.vTelefono);
                         command.Parameters.AddWithValue("@vCelular",entidad.vCelular);
                         command.Parameters.AddWithValue("@vCorreo",entidad.vCorreo);
                         command.Parameters.AddWithValue("@vNombreRepre",entidad.vNombreRepre);
                         command.Parameters.AddWithValue("@vApepatRepre",entidad.vApepatRepre);
                         command.Parameters.AddWithValue("@vApematRepre",entidad.vApematRepre);
                         command.Parameters.AddWithValue("@bSexo",entidad.bSexo);
                         command.Parameters.AddWithValue("@vDniRepre",entidad.vDniRepre);
                         command.Parameters.AddWithValue("@vCargoDesempeña",entidad.vCargoDesempeña);
                         command.Parameters.AddWithValue("@vTelefonoRepre",entidad.vTelefonoRepre);
                         command.Parameters.AddWithValue("@vCelularRepre",entidad.vCelularRepre);
                         command.Parameters.AddWithValue("@vCorreoRepre",entidad.vCorreoRepre);
                         command.Parameters.AddWithValue("@iCodConvocatoria",entidad.iCodConvocatoria);
                         
                         command.Parameters.AddWithValue("@vNombres",entidad.vNombres);
                         command.Parameters.AddWithValue("@vApepat",entidad.vApepat);
                         command.Parameters.AddWithValue("@vApemat",entidad.vApemat);
                         command.Parameters.AddWithValue("@bSexoExt",entidad.bSexoExt);
                         command.Parameters.AddWithValue("@dFechaNacimiento",entidad.dFechaNacimiento);
                         command.Parameters.AddWithValue("@vDni",entidad.vDni);
                         command.Parameters.AddWithValue("@vRuc",entidad.vRuc);
                         command.Parameters.AddWithValue("@vCorreoExt",entidad.vCorreoExt);
                         command.Parameters.AddWithValue("@vTelefonoExt",entidad.vTelefonoExt);
                         command.Parameters.AddWithValue("@vCelularExt",entidad.vCelularExt);
                         command.Parameters.AddWithValue("@vCodDepartamentoExt",entidad.vCodDepartamentoExt);
                         command.Parameters.AddWithValue("@vCodProvinciaExt",entidad.vCodProvinciaExt);
                         command.Parameters.AddWithValue("@vCodDistritoExt",entidad.vCodDistritoExt);
                         command.Parameters.AddWithValue("@vDomicilio",entidad.vDomicilio);
                         command.Parameters.AddWithValue("@iCodNivelInstruccion",entidad.iCodNivelInstruccion);
                         command.Parameters.AddWithValue("@iCodDirecGerencia",entidad.iCodDirecGerencia);
                         command.Parameters.AddWithValue("@iCodAgenciaAgraria",entidad.iCodAgenciaAgraria);
                        command.Parameters.AddWithValue("@vClave",entidad.vClave);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {

                                    Extensionista Extensionista = new Extensionista();

                                    Extensionista.iCodExtensionista = dr.GetInt32(dr.GetOrdinal("iCodExtensionista"));
                                    Entidad.iCodExtensionista = Extensionista.iCodExtensionista;

                                    Extensionista.vMensaje = dr.GetString(dr.GetOrdinal("vMensaje"));
                                    Entidad.vMensaje = Extensionista.vMensaje;
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

        public Extensionista Login(Extensionista entidad)
        {
            Extensionista Entidad = new Extensionista();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    entidad.vClave = GetSHA256(entidad.vClave);
                    using (SqlCommand command = new SqlCommand("[PA_Login]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                                                
                        command.Parameters.AddWithValue("@vDni", entidad.vDni);
                        command.Parameters.AddWithValue("@vClave", entidad.vClave);

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if(dr.Read())
                                {                                                                           
                                    Entidad.iCodExtensionista = dr.GetInt32(dr.GetOrdinal("iCodExtensionista"));      
                                    Entidad.vNombres = dr.GetString(dr.GetOrdinal("vNombres"));
                                    Entidad.vApepat = dr.GetString(dr.GetOrdinal("vApepat"));
                                    Entidad.vApemat = dr.GetString(dr.GetOrdinal("vApemat"));
                                    Entidad.dFechaUltimoAcceso = dr.GetString(dr.GetOrdinal("dFechaUltimoAcceso"));
                                    Entidad.iCodEmpresa = dr.GetInt32(dr.GetOrdinal("iCodEmpresa"));
                                    Entidad.vNombrePropuesta = dr.GetString(dr.GetOrdinal("vNombrePropuesta"));
                                    Entidad.iEnvio = dr.GetInt32(dr.GetOrdinal("iEnvio"));
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

        public Extensionista ListarExtensionistaPorCodigo(Extensionista entidad)
        {
            Extensionista Entidad = new Extensionista();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();
                                        
                    using (SqlCommand command = new SqlCommand("[PA_Listar_Extensionista_Por_Codigo]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iCodExtensionista", entidad.iCodExtensionista);
                        
                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    Entidad.iCodExtensionista = dr.GetInt32(dr.GetOrdinal("iCodExtensionista"));
                                    Entidad.iCodConvocatoria = dr.GetInt32(dr.GetOrdinal("iCodConvocatoria"));
                                    Entidad.iCodEmpresa = dr.GetInt32(dr.GetOrdinal("iCodEmpresa"));
                                    Entidad.vNombres = dr.GetString(dr.GetOrdinal("vNombres"));
                                    Entidad.vApepat = dr.GetString(dr.GetOrdinal("vApepat"));
                                    Entidad.vApemat = dr.GetString(dr.GetOrdinal("vApemat"));
                                    Entidad.bSexo = dr.GetInt32(dr.GetOrdinal("bSexo"))==1 ? true :false;
                                    Entidad.dFechaNacimiento = dr.GetString(dr.GetOrdinal("dFechaNacimiento"));
                                    Entidad.vDni = dr.GetString(dr.GetOrdinal("vDni"));
                                    Entidad.vRuc = dr.GetString(dr.GetOrdinal("vRuc"));
                                    Entidad.vCorreo = dr.GetString(dr.GetOrdinal("vCorreo"));
                                    Entidad.vTelefono = dr.GetString(dr.GetOrdinal("vTelefono"));
                                    Entidad.vCelular = dr.GetString(dr.GetOrdinal("vCelular"));
                                    Entidad.vCodDistrito = dr.GetString(dr.GetOrdinal("vCodDistrito"));
                                    Entidad.vDomicilio = dr.GetString(dr.GetOrdinal("vDomicilio"));
                                    Entidad.iCodNivelInstruccion = dr.GetInt32(dr.GetOrdinal("iCodNivelInstruccion"));
                                    Entidad.iCodDirecGerencia = dr.GetInt32(dr.GetOrdinal("iCodDirecGerencia"));
                                    Entidad.iCodAgenciaAgraria = dr.GetInt32(dr.GetOrdinal("iCodAgenciaAgraria"));
                                    Entidad.vNomDepartamento = dr.GetString(dr.GetOrdinal("vNomDepartamento"));
                                    Entidad.vNomProvincia = dr.GetString(dr.GetOrdinal("vNomProvincia"));
                                    Entidad.vNomDistrito = dr.GetString(dr.GetOrdinal("vNomDistrito"));
                                    Entidad.vNombrePropuesta = dr.GetString(dr.GetOrdinal("vNombrePropuesta"));
                                    Entidad.vRazonSocial = dr.GetString(dr.GetOrdinal("vRazonSocial"));
                                    Entidad.vNombreRepre = dr.GetString(dr.GetOrdinal("vNombreRepre"));
                                    Entidad.vDniRepre = dr.GetString(dr.GetOrdinal("vDniRepre"));
                                    Entidad.vRucEmpresa = dr.GetString(dr.GetOrdinal("vRucEmpresa"));
                                    Entidad.vDomicilioFiscal = dr.GetString(dr.GetOrdinal("vDomicilioFiscal"));
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

        public List<Menu> ObtenerMenu(Menu menu)
        {
            List<Menu> lista = new List<Menu>();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();

                    using (SqlCommand command = new SqlCommand("[PA_MenuAccesoPerfil]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@iTipoReg", menu.iTipoReg);                

                        using (SqlDataReader dr = command.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {                               
                                while (dr.Read())
                                {
                                    menu = new Menu();
                                    menu.iCodMenu = dr.GetInt32(dr.GetOrdinal("iCodMenu"));
                                    menu.vTitulo = dr.GetString(dr.GetOrdinal("vTitulo"));
                                    menu.vRuta = dr.GetString(dr.GetOrdinal("vRuta"));
                                    menu.iPadre = dr.GetInt32(dr.GetOrdinal("iPadre"));
                                    lista.Add(menu);

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
    }
}
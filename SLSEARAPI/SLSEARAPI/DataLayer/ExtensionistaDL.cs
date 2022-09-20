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
    public class ExtensionistaDL
    {
        public Extensionista InsertarExtensionista(Extensionista entidad)
        {
            Extensionista Entidad = new Extensionista();
            try
            {
                using (SqlConnection conection = new SqlConnection(ConfigurationManager.ConnectionStrings["cnx"].ConnectionString))
                {
                    conection.Open();


                    using (SqlCommand command = new SqlCommand("[PA_InsertarExtensionista]", conection))
                    {
                        command.CommandType = CommandType.StoredProcedure;


                        command.Parameters.AddWithValue("@icodExtensionista", entidad.iCodExtensionista);
                         command.Parameters.AddWithValue("@vRazonSocial",entidad.vRazonSocial           );
                         command.Parameters.AddWithValue("@vInicialesSiglas",entidad.vInicialesSiglas       );
                         command.Parameters.AddWithValue("@iCodProcedencia",entidad.iCodProcedencia        );
                         command.Parameters.AddWithValue("@iCodRegimen",entidad.iCodRegimen            );
                         command.Parameters.AddWithValue("@iCodTipoEntidad",entidad.iCodTipoEntidad        );
                         command.Parameters.AddWithValue("@vOtros",entidad.vOtros                 );
                         command.Parameters.AddWithValue("@vRucEmpresa",entidad.vRucEmpresa            );
                         command.Parameters.AddWithValue("@vDomicilioFiscal",entidad.vDomicilioFiscal       );
                         command.Parameters.AddWithValue("@vCodDepartamento",entidad.vCodDepartamento       );
                         command.Parameters.AddWithValue("@vCodProvincia",entidad.vCodProvincia          );
                         command.Parameters.AddWithValue("@vCodDistrito",entidad.vCodDistrito           );
                         command.Parameters.AddWithValue("@vTelefono",entidad.vTelefono              );
                         command.Parameters.AddWithValue("@vCelular",entidad.vCelular               );
                         command.Parameters.AddWithValue("@vCorreo",entidad.vCorreo                );
                         command.Parameters.AddWithValue("@vNombreRepre",entidad.vNombreRepre           );
                         command.Parameters.AddWithValue("@vApepatRepre",entidad.vApepatRepre           );
                         command.Parameters.AddWithValue("@vApematRepre",entidad.vApematRepre           );
                         command.Parameters.AddWithValue("@bSexo",entidad.bSexo                  );
                         command.Parameters.AddWithValue("@vDniRepre",entidad.vDniRepre              );
                         command.Parameters.AddWithValue("@vCargoDesempeña",entidad.vCargoDesempeña        );
                         command.Parameters.AddWithValue("@vTelefonoRepre",entidad.vTelefonoRepre         );
                         command.Parameters.AddWithValue("@vCelularRepre",entidad.vCelularRepre          );
                         command.Parameters.AddWithValue("@vCorreoRepre",entidad.vCorreoRepre           );
                         command.Parameters.AddWithValue("@iCodConvocatoria",entidad.iCodConvocatoria       );
                         command.Parameters.AddWithValue("@iCodEmpresa",entidad.iCodEmpresa            );
                         command.Parameters.AddWithValue("@vNombres",entidad.vNombres               );
                         command.Parameters.AddWithValue("@vApepat",entidad.vApepat                );
                         command.Parameters.AddWithValue("@vApemat",entidad.vApemat                );
                         command.Parameters.AddWithValue("@bSexoExt",entidad.bSexoExt               );
                         command.Parameters.AddWithValue("@dFechaNacimiento",entidad.dFechaNacimiento       );
                         command.Parameters.AddWithValue("@vDni",entidad.vDni                   );
                         command.Parameters.AddWithValue("@vRuc",entidad.vRuc                   );
                         command.Parameters.AddWithValue("@vCorreoExt",entidad.vCorreoExt             );
                         command.Parameters.AddWithValue("@vTelefonoExt",entidad.vTelefonoExt           );
                         command.Parameters.AddWithValue("@vCelularExt",entidad.vCelularExt            );
                         command.Parameters.AddWithValue("@vCodDepartamentoExt",entidad.vCodDepartamentoExt    );
                         command.Parameters.AddWithValue("@vCodProvinciaExt",entidad.vCodProvinciaExt       );
                         command.Parameters.AddWithValue("@vCodDistritoExt",entidad.vCodDistritoExt        );
                         command.Parameters.AddWithValue("@vDomicilio",entidad.vDomicilio             );
                         command.Parameters.AddWithValue("@iCodNivelInstruccion",entidad.iCodNivelInstruccion   );
                         command.Parameters.AddWithValue("@iCodDirecGerencia",entidad.iCodDirecGerencia      );
                         command.Parameters.AddWithValue("@iCodAgenciaAgraria",entidad.iCodAgenciaAgraria     );
                        command.Parameters.AddWithValue("@vClave",entidad.vClave                 );



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
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class Extensionista
    {
        [DataMember] public int iCodExtensionista { get; set; }
        [DataMember] public string  vRazonSocial              {get;set;}
      [DataMember] public string vInicialesSiglas          {get;set;}
      [DataMember] public int  iCodProcedencia             {get;set;}
      [DataMember] public int  iCodRegimen                 {get;set;}
      [DataMember] public int  iCodTipoEntidad             {get;set;}
      [DataMember] public string vOtros                    {get;set;}
      [DataMember] public string vRucEmpresa               {get;set;}
      [DataMember] public string vDomicilioFiscal          {get;set;}
      [DataMember] public string vCodDepartamento          {get;set;}
      [DataMember] public string vCodProvincia             {get;set;}
      [DataMember] public string vCodDistrito              {get;set;}
      [DataMember] public string vTelefono                 {get;set;}
      [DataMember] public string vCelular                  {get;set;}
      [DataMember] public string vCorreo                   {get;set;}
      [DataMember] public string vNombreRepre              {get;set;}
      [DataMember] public string vApepatRepre              {get;set;}
      [DataMember] public string vApematRepre              {get;set;}
      [DataMember] public bool  bSexo                     {get;set;}
      [DataMember] public string vDniRepre                 {get;set;}
      [DataMember] public string vCargoDesempeña           {get;set;}
      [DataMember] public string vTelefonoRepre            {get;set;}
      [DataMember] public string vCelularRepre             {get;set;}
      [DataMember] public string vCorreoRepre              {get;set;}
      //[DataMember] public int                            {get;set;}
      //[DataMember] public int                            {get;set;}
      [DataMember] public int  iCodConvocatoria          {get;set;}
      [DataMember] public int  iCodEmpresa               {get;set;}
      [DataMember] public string vNombres                  {get;set;}
      [DataMember] public string vApepat                   {get;set;}
      [DataMember] public string vApemat                   {get;set;}
      [DataMember] public bool bSexoExt                  {get;set;}
      [DataMember] public string dFechaNacimiento          {get;set;}
      [DataMember] public string vDni                      {get;set;}
      [DataMember] public string vRuc                      {get;set;}
      [DataMember] public string vCorreoExt                {get;set;}
      [DataMember] public string  vTelefonoExt              {get;set;}
      [DataMember] public string  vCelularExt               {get;set;}
      [DataMember] public string  vCodDepartamentoExt       {get;set;}
      [DataMember] public string  vCodProvinciaExt          {get;set;}
      [DataMember] public string  vCodDistritoExt           {get;set;}
      [DataMember] public string vDomicilio                {get;set;}
      [DataMember] public int  iCodNivelInstruccion      {get;set;}
      [DataMember] public int  iCodDirecGerencia         {get;set;}
      [DataMember] public int  iCodAgenciaAgraria        {get;set;}
      [DataMember] public string vClave { get; set; }
      [DataMember] public string vMensaje { get; set; }
      [DataMember] public string dFechaUltimoAcceso { get; set; }
      [DataMember]  public string vNomDepartamento { get; set; }
      [DataMember] public string vNomProvincia { get; set; }
      [DataMember] public string vNomDistrito { get; set; }
      [DataMember] public string vNombrePropuesta { get; set; }
        [DataMember]  public int iEnvio { get; set; }
    }
}
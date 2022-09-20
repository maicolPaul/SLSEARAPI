using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class FormacionAcademica
    {
        [DataMember] public int iCodFormacionAcademica   {get;set;}
      [DataMember] public int iCodNivel                  {get;set;}
      [DataMember] public string vCentroEstudios         {get;set;}
      [DataMember] public string vEspecialidad           {get;set;}
      [DataMember] public string vMensaje { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class NombreArchivo
    {
      [DataMember] public int  iCodNombreArchivo    {get;set;}
      [DataMember] public string vNombreArchivo { get; set; }
    }
}
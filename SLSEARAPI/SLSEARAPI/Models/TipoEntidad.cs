using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class TipoEntidad
    {
      [DataMember] public int   iCodTipoEntidad     {get;set;}
      [DataMember] public string vTipoEntidad { get; set; }
    }
}
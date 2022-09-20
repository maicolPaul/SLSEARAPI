using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class DireccionGerenciaRegionalAgraria
    {
        [DataMember] public int iCodDirecGerencia      {get;set;}
        [DataMember] public string vNombre { get; set; }
    }
}
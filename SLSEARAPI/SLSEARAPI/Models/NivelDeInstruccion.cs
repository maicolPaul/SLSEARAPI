using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class NivelDeInstruccion
    {
       [DataMember] public int iCodNivelInstruccion     {get;set;}
       [DataMember] public string vDescripcion { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class Procedencia
    {
        [DataMember] public int iCodProcedencia   {get;set;}
        [DataMember] public string vProcedencia { get; set; }
    }
}
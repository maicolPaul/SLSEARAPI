using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class Requisitos
    {
       [DataMember] public int iCodRequisito       {get;set;}
       [DataMember] public string vRequisito { get; set; }
    }
}
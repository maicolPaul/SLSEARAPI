using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{

    [DataContract]
    public class AgenciaAgraria
    {
       [DataMember] public int  iCodAgenciaAgraria       {get;set;}
       [DataMember] public string vAgencia { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class Regimen
    {
      [DataMember] public int iCodRegimen  {get;set;}
      [DataMember] public string vRegimen { get; set; }
    }
}
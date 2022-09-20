using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class ListaChequeoRequisitos
    {
      [DataMember] public int  iCodExtensionista     {get;set;}
      [DataMember] public int  iCodRequisito         {get;set;}
      [DataMember] public bool  bCumple              {get;set;}
      [DataMember] public string vMensaje { get; set; }
        [DataMember] public int iCodListaChequeo { get; set; }
    }
}
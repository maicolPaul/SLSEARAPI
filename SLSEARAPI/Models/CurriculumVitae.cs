using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class CurriculumVitae
    {
      [DataMember] public int iCodExtensionista          {get;set;}
      [DataMember] public int iCodFormacionAcademica     {get;set;}
      [DataMember] public int iCodExperiencia            { get; set; }
      [DataMember] public int  iCodCurriculumVitae       {get;set;}
      [DataMember] public string vMensaje                { get; set; }
    }
}
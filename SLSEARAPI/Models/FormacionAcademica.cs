using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class FormacionAcademica
    {
      [DataMember] public int iCodFormacionAcademica   {get;set;}
      [DataMember] public int iCodNivel                  {get;set;}
      [DataMember]  public string descripcionnivel { get; set; }
      [DataMember] public string vCentroEstudios         {get;set;}
      [DataMember] public string vEspecialidad           {get;set;}
      [DataMember] public string vMensaje { get; set; }
      [DataMember] public int iCodExtensionista { get; set; }
      [DataMember] public int iCodCurriculumVitae { get; set; }

        [DataMember] public int piPageSize { get; set; }
        [DataMember]  public int piCurrentPage { get; set; }
        [DataMember]  public string pvSortColumn { get; set; }
        [DataMember]  public string pvSortOrder { get; set; }
        [DataMember]  public int totalRegistros { get; set; }
        [DataMember]  public int totalPaginas { get; set; }

        [DataMember]  public int paginaActual { get; set; }

        [DataMember]  public int iOpcion { get; set; }

    }
}
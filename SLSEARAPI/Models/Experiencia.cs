using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class Experiencia
    {
      [DataMember] public string vNombreEntidad             {get;set;}
      [DataMember] public string vCargoServicio             {get;set;}
      [DataMember] public string vActividades               { get; set; }
      [DataMember] public string vProductoServicio          { get; set; }
      [DataMember] public string dFechaInicio               {get;set;}
      [DataMember] public string dFechaFin                  {get;set;}
      [DataMember] public string vRutaArchivoConstancia { get; set; }
        [DataMember] public int iCodExperiencia { get; set; }
        [DataMember] public int iCodCurriculumVitae { get; set; }
        [DataMember] public string vMensaje { get; set; }

        [DataMember]
        public int iTipoExperiencia { get; set; }

        //SUBIR ARCHIVO MODEL
        [DataMember]
        public byte[] file { get; set; }

        [DataMember]
        public string path { get; set; }

        [DataMember]
        public string mineType { get; set; }

        [DataMember]
        public string mensaje { get; set; }
        [DataMember]
        public string[] fileExtensiones { get; set; }
        [DataMember]
        public bool validation { get; set; }
        [DataMember]
        public List<string> fileNames { get; set; }
        [DataMember]
        public string encode64 { get; set; }
        [DataMember]
        public int piPageSize { get; set; }
        [DataMember]
        public int piCurrentPage { get; set; }
        [DataMember]
        public string pvSortColumn { get; set; }
        [DataMember]
        public string pvSortOrder { get; set; }
        [DataMember]
        public int totalRegistros { get; set; }
        [DataMember]
        public int totalPaginas { get; set; }
        [DataMember]
        public int paginaActual { get; set; }
        [DataMember]
        public int meses { get; set; }
        [DataMember]
        public int totalmeses { get; set; }
        [DataMember]
        public int iOpcion { get; set; }
    }

}
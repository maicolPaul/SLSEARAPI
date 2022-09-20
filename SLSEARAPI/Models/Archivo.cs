using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace SLSEARAPI.Models
{
    [DataContract]
    public class Archivo
    {
       [DataMember] public int icodExtensionista     {get;set;}
       [DataMember] public int iCodNombreArchivo     {get;set;}
       [DataMember] public string vRutaArchivo { get; set; }

       [DataMember] public int iCodArchivos { get; set; }

       [DataMember] public string vMensaje { get; set; }

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
    }
}
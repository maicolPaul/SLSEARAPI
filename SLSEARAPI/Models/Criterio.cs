using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class Criterio
    {
        public int iCodCriterio { get; set; }

        public string vDescripcion { get; set; }

        public int iRecordCount { get; set; }

        public int iPageCount { get; set; }

        public int iCurrentPage { get; set; }

        public int iCodRubro { get; set; }

        public int piPageSize { get; set; }

        public int piCurrentPage { get; set; }

        public string pvSortColumn { get; set; }

        public string pvSortOrder { get; set; }

        public int iCodSuperCab { get; set; }

        public int iCodCalificacion { get; set; }

        public string vFundamento { get; set; }

        public string vDescripcionCal { get; set; }

    }
}
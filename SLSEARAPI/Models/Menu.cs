using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SLSEARAPI.Models
{
    public class Menu
    {
        public int iCodMenu { get; set; }

        public string vTitulo { get; set; }

        public string vRuta { get; set; }

        public int iTipoReg { get; set; }

        public int iPadre { get; set; }

        public List<Menu> subitems { get; set; }
    }
}
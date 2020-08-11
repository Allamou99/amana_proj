using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
namespace bamEpplus
{
        abstract class Ligne
        {
            public String Identificateur { get; set; }// identifier le type de ligne
            public ICollection<Donnee> liste { get; set; } // la liste des elements de la ligne 
            public int nombre_element { get; set; } // le nombre des elements 
            public char delimiteur { get; set; } // delimeteur pour specifie le type de gestion 
            //public abstract string lire_ligne(String ligne);
            public abstract void ecrire_ligne(int index, ExcelWorksheet worksheet); // la fonction pour ecrire les données sous excel 

        }
}

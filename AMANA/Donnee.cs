using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace bamEpplus
{
    class Donnee
    {
            public string libelle { get; set; } // Le champs a afficher c'est on a une table 
            public string contenu { get; set; } // contenu de la table 
            public string type { get; set; } // type soit delimeteur ou index 
            public int index_debut { get; set; } // index de debut de lecture donnes c'est le type est index 
            public int index_fin { get; set; }   // index de fin de lecture donnes c'est le type est index 
            public int longueur { get; set; }
            public int index;                    // index de la cellule dans la table 
            public string cellule_debut { get; set; } // cellule debut 
            public string cellule_fin { get; set; }   // cellule fin 
            public double width;                      // la grandeur 

   

            public String affiche()
            {
                return "entete est :" + this.libelle + "\t et contenu est :" + this.contenu + "\t INDB: " + this.index_debut + "\t INDEX fin : "+this.index_fin + "\t Longueur : "+this.longueur;
            }

    }
}
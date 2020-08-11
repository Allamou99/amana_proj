using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Xml.Linq;


namespace bamEpplus
{
    class Shema_TRA
    {
        public String name;
        public String identificateur;
        public String abr;
        public String chemin_shema;
        public Dictionary<string, Entete> entetes = new   System.Collections.Generic.Dictionary<string,Entete>();
        public Villes lesVilles;
        public int nombre_lignes;

        public int init(String chemin_shema)  
        {
                 Console.WriteLine("Lecture des villes ........");
                 lesVilles= Villes.getInstance();
                 //Console.WriteLine("Fin de lecture");                 
                 String fichier_config = "";    
                 fichier_config = chemin_shema + @"\CONFG_shema.xml";
                 XElement xelement = XElement.Load(@fichier_config);
                 var lines = from line in xelement.Elements("ENTETES").Elements()
                                  select line;
                 Console.WriteLine("Lecture des parametres de shera TRA ........");
                 foreach (var line in lines)
                 {
                        String nombre_element_string = (String)line.Attribute("TAILLE");
                        int nombre_element_node = Convert.ToInt32(nombre_element_string);
                        String nom_nelement_node=(string)line.Attribute("ID");
                        String node_type = (string)line.Attribute("TYPE");
                        Entete entete = new Entete { nombre_element = nombre_element_node, Identificateur = nom_nelement_node, type=node_type, delimiteur='i' };
                        entete.lire_XML(line);
                        //Console.WriteLine(entete);/*****************************/
                        entetes.Add(nom_nelement_node, entete);
                 }

                // Console.WriteLine("Fin de lecture"); 
                  return 1;
        
        }
        public void affiche() {

                Console.WriteLine("------------------------------------------");
                foreach (KeyValuePair<string, Entete> entete in entetes)
                {
                    Console.WriteLine("Key = {0}", entete.Key );           
                    Console.WriteLine("***********");
                    entete.Value.affiche();
                }

                Console.WriteLine("------------------------------------------");
        
        }




    }
}

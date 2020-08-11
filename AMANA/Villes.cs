using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Xml.Linq;
using System.IO;
using OfficeOpenXml;

namespace bamEpplus
{
    public sealed class Villes
    {
        public String chemin_app = "";
        public Dictionary<string, string> Listes_Villes = new System.Collections.Generic.Dictionary<string, string>();

         private static Villes villes = new Villes();


        private Villes()
        {
            initialisation();
            // TODO: Complete member initialization
        }
        public static Villes getInstance()
        {

            return villes;
        }
        private Villes initialisation()
        {
            String codeville = "";
            String localite="";
            String fichier_config = "";
            try
            {
                chemin_app = ConfigurationManager.AppSettings["VILLES"];
                fichier_config = chemin_app + @"\Villes.xlsx";
                FileInfo existingFile = new FileInfo(fichier_config);
                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                      
                    int i = 2;
                    while (worksheet.Cells[i, 1].Value !=null)
                    {
                           
                        codeville = worksheet.Cells[i, 1].Value.ToString();
                        localite = worksheet.Cells[i, 2].Value.ToString();

                        if (!(codeville.CompareTo("") == 0))
                        {
                            Listes_Villes.Add(localite, codeville);
                        }
                        else
                        {
                            break;
                        }
                        //Console.WriteLine("code ville :"+ codeville +" localité :" +localite);
                        //Console.WriteLine("********");
                        i++;

                    }

                } // fermer le package;
            }catch(Exception ex){
                Console.WriteLine(ex.ToString());
                
            }

            return this;
        }
        public void affiche()
        {
            foreach (KeyValuePair<string, string> ville in this.Listes_Villes)
            {
                
                Console.WriteLine("code ville est :" + ville.Key + "\t localité ville est :" + ville.Value);
              
                
                    
            }


        }

        public string chercher_ville(string localite) 
        {
                    String codeVille="";
                    bool resultat=Listes_Villes.TryGetValue(localite, out codeVille);

                    if (resultat == true) 
                        { return codeVille; }
                    else
                    {
                        String Sanslocalite = "";
                        Listes_Villes.TryGetValue("SANSLOCALITE", out Sanslocalite);

                        return Sanslocalite;
                    }               
        }

          
    }
}

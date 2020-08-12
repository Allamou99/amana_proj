using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Configuration;

namespace bamEpplus
{
    public static class Utilitaire
    {


            public static String calcule_poid(String poids){
                        int partie_entiere = Convert.ToInt32(poids.Substring(0, 6));
                        int partie_decimale = Convert.ToInt32(poids.Substring(6, 3));
                        if (partie_decimale == 0) return partie_entiere+"";
                        return partie_entiere + "." + partie_decimale;
            }
            public static String calcule_double(Double chaine2, int lng_entriere, int lng_dec)
            {
                String chaine = chaine2.ToString();
                int partie_entiere = Convert.ToInt32(chaine.Substring(0, lng_entriere));
                int partie_decimale = Convert.ToInt32(chaine.Substring(lng_entriere, lng_dec));
                if (partie_decimale == 0) return partie_entiere + "";
                return partie_entiere + "." + partie_decimale;
            }

            public static String supprimer_espace(String valuer ) {

                         return valuer.Trim();        
        
            }

            public static string num_fichier(Parametrage parametre)
            {
                        
                        int max_base=0;
                        int temp = 0;
                        String nom_fichier = "";
                        String base_nom_fichier = parametre.prefixe + "_" + DateTime.Now.ToString("ddMMyyyy");
                        var fichiers = Directory.GetFiles(@parametre.chemin_genration_excel);
                        foreach (String fichier in fichiers)
                        {
                            var mots = fichier.Split('\\');


                            if ((mots[mots.Length - 1].Substring(0, 13).CompareTo(base_nom_fichier)) == 0)
                            {
                                temp = Convert.ToInt32(mots[mots.Length - 1].Substring(14, 3));
                                if (temp > max_base) 
                                {
                                    max_base = temp; 
                                }                        
                            }
                        }
                        fichiers = Directory.GetFiles(@parametre.chemin_archive_excel);
                        foreach (String fichier in fichiers)
                        {
                            var mots = fichier.Split('\\');


                            if ((mots[mots.Length - 1].Substring(0, 13).CompareTo(base_nom_fichier)) == 0)
                            {
                                temp = Convert.ToInt32(mots[mots.Length - 1].Substring(14, 3));
                                if (temp > max_base) { max_base = temp; }

                            }
                        }
                        max_base = max_base + 1;
                        if (max_base < 10) { nom_fichier = base_nom_fichier + "_00" + max_base+ ".xlsx";  }
                        else
                        {
                            if (max_base < 100) { nom_fichier = base_nom_fichier + "_0" + max_base + ".xlsx"; }
                            else { nom_fichier = base_nom_fichier + "_" + max_base + ".xlsx"; }      
                        }
                        return nom_fichier;
           }


            public static bool fichier_trace(String chemin, String trace)
            {
                String nom_fichier = "";
                bool existe;
                //C:\debugage\AMANA_131\trace
                DateTime localDate = DateTime.Now;
                chemin = ConfigurationManager.AppSettings["CHEMINTRACE"];
                nom_fichier = chemin + @"\Amana_131_" + localDate.Year + "" + localDate.Month + "" + localDate.Day + "_000001.log";
                existe = File.Exists(@nom_fichier);
                if (!existe)
                {
                    var fichiers = Directory.GetFiles(@chemin);

                    foreach (String fichier in fichiers)
                    {
                        var mots = fichier.Split('\\');
                        String destFile = ConfigurationManager.AppSettings["CHEMINTRACEOLD"] + @"\" + mots[mots.Length - 1];
                        
                        System.IO.File.Copy(fichier, destFile, true);
                        System.IO.File.Delete(fichier);
                    }
                }

                using (StreamWriter sw = File.AppendText(@nom_fichier))
                {
                    sw.WriteLine("{0} AMANA_131 : {1}", DateTime.Now.ToLongTimeString(), trace);
                }	

               // System.IO.File.AppendAllLines(@nom_fichier, trace);
                return true;
            
            }
    }
}
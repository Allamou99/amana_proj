using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AMANA
{
    class Old_Generation
    {
    }
}


/*  StreamWriter monStreamWriterERR = new StreamWriter(@chemin_complet_erreur);

  monStreamWriterERR.WriteLine("Géneration de fichier excel pour fichier : " + fichier );
  monStreamWriterERR.Close();*/
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.IO;
//using OfficeOpenXml;
//using OfficeOpenXml.Style;

//namespace bamEpplus
//{
//    class Generation
//    {

//        public String identificateur;

//        public String fichier_traite;
//        public Shema_TRA shema;
//        public Dictionary<string, ICollection<LigneContenu>> liste_lignes_fichier = new System.Collections.Generic.Dictionary<string, ICollection<LigneContenu>>();
//        public Parametrage parametre;

//        // Constructeur de la classe 
//        public Generation(Shema_TRA shema, String identificateur, Parametrage parametre)
//        {
//                this.identificateur = identificateur;
//                this.shema = shema;
//                this.parametre = parametre;
//        }

//        /**
//         * fonction permet de lire le fichier data passe en parametre et de construire 
//         * la liste des lignes qui sert a la generation de fichier excel 
//         */

//        public bool lire_PLAT(String fichier)
//        {

//            try
//            {
//                StreamReader monStreamReader = new StreamReader(@fichier);
//                String ligne = "";
//                String identificateur_ligne;

//                int i = 0;
//                while ((ligne = monStreamReader.ReadLine()) != null)
//                {
//                    identificateur_ligne = ligne.Substring(0, 6);
//                    var lignecontenu = new LigneContenu { Identificateur = identificateur_ligne };
//                    Entete entete;
//                    shema.entetes.TryGetValue(identificateur_ligne, out entete);
//                    lignecontenu.lire_ligne(ligne, entete, i);
//                    ICollection<LigneContenu> Laligne_fichier;
//                    if (!this.liste_lignes_fichier.ContainsKey(identificateur_ligne))
//                    {  /**
//                        * le dectionnaire ne contient aucune occurrence de la cle   
//                         *donc on creer une nouvelle liste de valuer apré on ajoute la valuer a la liste ensuite on ajoute le couple
//                         *au dictionnaire 
//                        **/
//                        Laligne_fichier = new List<LigneContenu>();
//                        Laligne_fichier.Add(lignecontenu);
//                        this.liste_lignes_fichier.Add(identificateur_ligne, Laligne_fichier);
//                    }
//                    else
//                    { // la cle deja existe donc on ajoute a liste des valeurs de la cle la nouvelle valuer  
//                        this.liste_lignes_fichier.TryGetValue(identificateur_ligne, out Laligne_fichier);
//                        Laligne_fichier.Add(lignecontenu);
//                    }
//                    i++;

//                }
//                //  Console.ReadKey();
//                monStreamReader.Close();// fermeture de fichier 
//                fichier_traite = fichier;// stock" le nom de fichier traité 
//                return true;
//            }catch(Exception ex){

//                Console.WriteLine(ex.ToString());
//                return false;
//            }
                
//        }





//        public bool ecrire_excel()
//        {
//            int index;
//            String produit="27";//16 par default
//            String num_envoi="";//check digit ==> colis 
//            String destination="";// code de la ville du client
//            String poids="";// poid de colis
//            String CRBT_CCP="";// contre-rmbrssmnt
//            String ville="";// ville du client 
//            String destinataire="";// Raison sociale du client
//            String code_lient = "";
//            Console.WriteLine("Géneration de fichier excel .....");
//            try {   
//                    // nom template: par exemple a mentionné dans la configue  Import en masse.xlsx
//                using (ExcelPackage p = new ExcelPackage(new FileInfo(@parametre.chemin_template + @"\" + parametre.nom_fichier_template), true))
//                    {
//                        p.Workbook.Properties.Author = "YON-ASIS-MEA";
//                        p.Workbook.Properties.Title = "titre de fichier";
//                        p.Workbook.Properties.Company = "BAM";
//                        var mots = this.fichier_traite.Split('\\');
//                        ExcelWorksheet worksheet = p.Workbook.Worksheets["Sheet1"];
//                        index = 2;

//                        /** Début de déclaration des lignes d'expedition Standard */
//                        ICollection<LigneContenu> Laligne_fichier_EEESTD;
//                        ICollection<LigneContenu> Laligne_fichier_ECESTD;
//                        ICollection<LigneContenu> Laligne_fichier_EBESTD;
//                        ICollection<LigneContenu> Laligne_fichier_EBLSTD;
//                        ICollection<LigneContenu> Laligne_fichier_EBCSTD;
//                        this.liste_lignes_fichier.TryGetValue("EEESTD", out Laligne_fichier_EEESTD);
//                        this.liste_lignes_fichier.TryGetValue("ECESTD", out Laligne_fichier_ECESTD);
//                        this.liste_lignes_fichier.TryGetValue("EBESTD", out Laligne_fichier_EBESTD);
//                        this.liste_lignes_fichier.TryGetValue("EBLSTD", out Laligne_fichier_EBLSTD);
//                        this.liste_lignes_fichier.TryGetValue("EBCSTD", out Laligne_fichier_EBCSTD);
//                        /** Fin de déclaration des lignes Standard **/
//                        /** Début de déclaration des lignes d'expedition spécifique */
//                        //ICollection<LigneContenu> Laligne_fichier_EBLSPC;
//                        //ICollection<LigneContenu> Laligne_fichier_EBCSPC;
//                        //bool EBLSPC_resultat = this.liste_lignes_fichier.TryGetValue("EBLSPC", out Laligne_fichier_EBLSPC);
//                        //bool EBCSPC_resultat = this.liste_lignes_fichier.TryGetValue("EBCSPC", out Laligne_fichier_EBCSPC);
//                        /** Fin de déclaration des lignes spécifiques **/                       
   
//                        foreach (LigneContenu entete in Laligne_fichier_EBCSTD)
//                        {
//                            String chaine_code_client = "";
//                            foreach (Donnee donne in entete.liste)
//                            {
//                                switch (donne.libelle)
//                                {
//                                    case "CABINFOTRANSPORTEUR": num_envoi = Utilitaire.supprimer_espace(donne.contenu);
//                                        break;
//                                    case "POID": poids = Utilitaire.calcule_poid(donne.contenu);
//                                        break;
//                                    case "CODECLIENT": code_lient = Utilitaire.supprimer_espace(donne.contenu);
//                                        break;

//                                    default: break;
//                                }

//                            }
//                            foreach (LigneContenu laligne in Laligne_fichier_ECESTD)
//                            {

//                                //laligne.liste.Contains(new Donnee { libelle =parametre, contenu =  valuer});
//                                foreach (Donnee donne in laligne.liste)
//                                {
//                                    switch (donne.libelle)
//                                    {
//                                        case "CODECLIENT": chaine_code_client = Utilitaire.supprimer_espace(donne.contenu);
//                                            break;
//                                        case "RAISONSOCIALE": destinataire = Utilitaire.supprimer_espace(donne.contenu);
//                                            break;
//                                        case "VILLE": ville = Utilitaire.supprimer_espace(donne.contenu);
//                                            break;
//                                        default: break;
//                                    }

//                                }
//                                if (chaine_code_client.CompareTo(code_lient) == 0) break;

//                            }

//                            //destinataire = Utilitaire.chercher_donne(Laligne_fichier_ECESTD, "RAISONSOCIALE","CODECLIENT",code_lient);// RAISONSOCIALE      
//                            //ville = Utilitaire.chercher_donne(Laligne_fichier_ECESTD, "VILLE", "CODECLIENT", code_lient);//VILLE
//                            destination = shema.lesVilles.chercher_ville(ville); 

//                            for (int i = 1; i < 10; i++)
//                            {
//                                worksheet.Cells[index, i].StyleID = worksheet.Cells[2, i].StyleID;
//                                worksheet.Cells[index, i].Formula = worksheet.Cells[2, i].Formula;
                           


//                            }
//                            worksheet.Cells[index, 1].Value = produit;
//                            worksheet.Cells[index, 2].Value = num_envoi;
//                            worksheet.Cells[index, 3].Value = destination;
//                            worksheet.Cells[index, 4].Value = poids;
//                            worksheet.Cells[index, 5].Value = "";
//                            worksheet.Cells[index, 6].Value =  "";
//                            worksheet.Cells[index, 7].Value = "";
//                            worksheet.Cells[index, 8].Value = ville;
//                            worksheet.Cells[index, 9].Value = destinataire;

//                                index++;

//                        }
//                            Console.WriteLine("Fichier excel généré");
//                            String nom_fichier=Utilitaire.num_fichier(parametre);
//                            String chemin_complet_fichier = @parametre.chemin_genration_excel + @"\" + nom_fichier;
//                            Byte[] bin = p.GetAsByteArray();
//                            File.WriteAllBytes(@chemin_complet_fichier, bin);
//                            Console.WriteLine("Enregistrer Fichier excel ");

                            
//                    }
//                }catch(Exception ex)
//                {
//                            Console.WriteLine(ex.ToString());

//                }
//                return true;
//        }




//        public void affiche()
//        {
//                    foreach (KeyValuePair<string, ICollection<LigneContenu>> ligne_fichier in this.liste_lignes_fichier)
//                    {
//                        Console.WriteLine("*******************Key**********************");
//                        Console.WriteLine("Identificateur de la ligne est :" + ligne_fichier.Key);
//                        Console.WriteLine("*******************Value**********************");
//                        foreach (LigneContenu contenu in ligne_fichier.Value)
//                        {
//                            contenu.affiche();
//                        }
//                        Console.WriteLine("-----------------------------------------");
//                        Console.WriteLine();
//                    }
        
        
//        }
//    }
//}

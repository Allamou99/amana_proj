using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace bamEpplus
{
    class Generation
    {

        public String identificateur;
        public String fichier_traite;
        public Shema_TRA shema;
        public Dictionary<string, ICollection<LigneContenu>> liste_lignes_fichier = new System.Collections.Generic.Dictionary<string, ICollection<LigneContenu>>();
        public Parametrage parametre;

        // Constructeur de la classe 
        public Generation(Shema_TRA shema, String identificateur, Parametrage parametre)
        {
                this.identificateur = identificateur;
                this.shema = shema;
                this.parametre = parametre;
        }

        /**
         * fonction permet de lire le fichier data passe en parametre et de construire 
         * la liste des lignes qui sert a la generation de fichier excel 
         */

        public bool lire_PLAT(String fichier)
        {
            try
            {
                StreamReader monStreamReader = new StreamReader(@fichier);
                String ligne = "";
                String identificateur_ligne;

                int i = 0;
                while ((ligne = monStreamReader.ReadLine()) != null)
                {
                    identificateur_ligne = ligne.Substring(0, 6);
                    var lignecontenu = new LigneContenu { Identificateur = identificateur_ligne };
                    Entete entete;
                    shema.entetes.TryGetValue(identificateur_ligne, out entete);
                    lignecontenu.lire_ligne(ligne, entete, i);
                    ICollection<LigneContenu> Laligne_fichier;
                    if (!this.liste_lignes_fichier.ContainsKey(identificateur_ligne))
                    {  /**
                        * le dectionnaire ne contient aucune occurrence de la cle   
                         *donc on creer une nouvelle liste de valuer apré on ajoute la valuer a la liste ensuite on ajoute le couple
                         *au dictionnaire 
                        **/
                        Laligne_fichier = new List<LigneContenu>();
                        Laligne_fichier.Add(lignecontenu);
                        this.liste_lignes_fichier.Add(identificateur_ligne, Laligne_fichier);
                    }
                    else
                    { // la cle deja existe donc on ajoute a liste des valeurs de la cle la nouvelle valuer  
                        this.liste_lignes_fichier.TryGetValue(identificateur_ligne, out Laligne_fichier);
                        Laligne_fichier.Add(lignecontenu);
                    }
                    i++;

                }
                //  Console.ReadKey();
                monStreamReader.Close();// fermeture de fichier 
                fichier_traite = fichier;// stock" le nom de fichier traité 
                return true;
            }catch(Exception ex){

                Console.WriteLine(ex.ToString());
                return false;
            }
                
        }





        public bool ecrire_excel()
        {
            Console.WriteLine("we're in");
            //A.ISBAINE@poste.ma
            int index;
            String num_envoi = "";//Check digit ==> colis 
            String destination = "";// Code de la ville du client
            String poids = "";// Poid de colis
            String ville = "";// Ville du client 
            String destinataire = "";// Raison sociale du client
            //String code_lient = "";
            String valeur_declaree = ""; //valeur du montant declarée de la marchandise
            String expediteur_sms = "";// Numero pour envoyer SMS a l'éxpediteur
            String client_sms = "";// Numero pour envoyer SMS au client
            int contre_remboursement = 0;// Contre-rmbrssmnt oui ou non 
            int retour_accuse = 0; // Retour du document accusé de réception
            String valeur_montant_CRBT = "";//Valeur du montant contre-rmbrssmnt
            int fragile = 0;//La fragilité des produits de colis
            String adresse_postal;
            String adresse1 = "", adresse2 = "", adresse3 = "", adresse4 = "";
            Double code_postal=0;

            //ICollection<String> trace = new List<String>();
            //DateTime localDate = DateTime.Now;
            //String chemin_complet_erreur = parametre.chemin_archive_data + @"\trace\Amana_131_" + localDate.Year + "" + localDate.Month + "" + localDate.Day + "_000001.log";
           
            try
            {
                // nom template: par exemple a mentionné dans la configue  Import en masse.xlsx
                using (ExcelPackage p = new ExcelPackage(new FileInfo(@parametre.chemin_template + @"\" + parametre.nom_fichier_template), true))
                {
                    p.Workbook.Properties.Author = "YON-ASIS-MEA";
                    p.Workbook.Properties.Title = "titre de fichier";
                    p.Workbook.Properties.Company = "BAM";
                    var mots = this.fichier_traite.Split('\\');
                    ExcelWorksheet worksheet = p.Workbook.Worksheets["Sheet1"];
                    index = 2;
                    
                    try
                    {
                        //ExcelPackage p = new ExcelPackage(new FileInfo(@"C:\Users\soufiane\Documents\exel_to_xml.xlsx"), true) ;
                        FileInfo newFile = new FileInfo(@"C:\Users\soufiane\Documents\Classeur3.xlsx");
                        ExcelPackage p2 = new ExcelPackage(newFile);
                        
                        {
                            ExcelWorksheet worksheet2 = p2.Workbook.Worksheets["Feuil1"];
                            int i = 5;
                            int j = 8;
                            while (worksheet2.Cells[j, i].Value != null)
                            {
                                /*while (worksheet2.Cells[j, i].Value != null)
                                {
                                    Console.WriteLine(worksheet2.Cells[j, i].Value);
                                    i++;
                                }
                                j++;
                                i = 5;
                                */
                                switch(worksheet2.Cells[j,i].Value)
                                {
                                    case "NUMEROEXP": num_envoi = (string)worksheet2.Cells[j+1, i].Value;
                                        break;
                                    case "ADRESSEPOSTAL": code_postal = (double)worksheet2.Cells[j + 1, i].Value;
                                        break;
                                    case "VILLE": ville = (string)worksheet2.Cells[j + 1, i].Value;
                                        break;
                                    case "POID":
                                        poids = Utilitaire.calcule_double((double)worksheet2.Cells[j + 1, i].Value, 2, 3);
                                        break;
                                    case "CONTREREMBOURSEMENT": contre_remboursement = (int)worksheet2.Cells[j + 1, i].Value;
                                        break;
                                
                                }
                                i++;

                            }
                            }
                            Console.WriteLine("done");
                        
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("voici l'exeption : " + ex);
                    }

                    destination = shema.lesVilles.chercher_ville(ville);
                    adresse_postal = adresse1 + " " + adresse2 + " " + adresse3 + " " + adresse4 + " ";

                        for (int i = 1; i < 14; i++)
                        {
                            worksheet.Cells[index, i].StyleID = worksheet.Cells[2, i].StyleID;
                            worksheet.Cells[index, i].Formula = worksheet.Cells[2, i].Formula;
                        }
                        worksheet.Cells[index, 1].Value = parametre.code_produit_SMI;
                        worksheet.Cells[index, 2].Value = num_envoi;
                        worksheet.Cells[index, 3].Value = destination;
                        worksheet.Cells[index, 4].Value = poids;
                        if (contre_remboursement == 1)
                        {
                            worksheet.Cells[index, 5].Value = valeur_montant_CRBT;
                            worksheet.Cells[index, 6].Value = "";
                        }
                        else if (contre_remboursement == 2)
                        {
                            worksheet.Cells[index, 5].Value = "";
                            worksheet.Cells[index, 6].Value = valeur_montant_CRBT;
                        }
                        else
                        {
                            worksheet.Cells[index, 5].Value = "";
                            worksheet.Cells[index, 6].Value = "";
                        }

                        worksheet.Cells[index, 7].Value = code_postal;
                        worksheet.Cells[index, 8].Value = valeur_declaree;
                        worksheet.Cells[index, 9].Value = retour_accuse;
                        worksheet.Cells[index, 10].Value = expediteur_sms;
                        worksheet.Cells[index, 11].Value = client_sms;
                        worksheet.Cells[index, 12].Value = fragile;
                        worksheet.Cells[index, 13].Value = ville;
                        index++;

                
                // Console.WriteLine("Fichier excel généré");
                String nom_fichier = Utilitaire.num_fichier(parametre);
                    String chemin_complet_fichier = @parametre.chemin_genration_excel + @"\" + nom_fichier;
                    Console.WriteLine(chemin_complet_fichier);
                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(@chemin_complet_fichier, bin);
                    exel_to_txt_two(chemin_complet_fichier);


                Utilitaire.fichier_trace(parametre.chemin_archive_data, "Fichier excel généré et enregistrer " + chemin_complet_fichier);
                    //trace.Add("Fichier excel généré et enregistrer " + chemin_complet_fichier);
                    //Console.WriteLine("Enregistrer Fichier excel ");

                }
            }
            // System.IO.File.AppendAllLines(@chemin_complet_erreur, trace);
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                Utilitaire.fichier_trace(parametre.chemin_archive_data, "EXCEPTION de fichier excel ....." + ex.ToString());
                //trace.Add("EXCEPTION de fichier excel ....." + ex.ToString());
                //System.IO.File.AppendAllLines(@chemin_complet_erreur, trace);
                
            }
            return true;
        }
            
            
        
        
        
        static void exel_to_txt_two(String path)
        {
            try
            {
                //ExcelPackage p = new ExcelPackage(new FileInfo(@"C:\Users\soufiane\Documents\exel_to_xml.xlsx"), true) ;
                FileInfo newFile = new FileInfo(path);
                ExcelPackage p = new ExcelPackage(newFile);
                {
                    string fileName = @"C:\test\AMANA_131\Final\stg.txt";
                    if (File.Exists(fileName))
                    {
                        File.Delete(fileName);
                        Console.WriteLine("yo");
                    }
                    FileStream fs = File.Create(fileName);
                    fs.Close();
                    StreamWriter sw = new StreamWriter(fileName, true);
                    ExcelWorksheet worksheet = p.Workbook.Worksheets["Sheet1"];
                    int i = 1;
                    int j = 1;
                    while (worksheet.Cells[j, i].Value != null)
                    {
                        Console.WriteLine("too");
                        while (worksheet.Cells[j, i].Value != null)
                        {
                            sw.Write(worksheet.Cells[j, i].Value + "          ");
                            i++;
                        }
                        j++;
                        i = 1;
                        sw.WriteLine("");

                    }
                    sw.Close();
                    
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("voici l'exeption : " + ex);
            }

        }




        public void affiche()
        {
                    foreach (KeyValuePair<string, ICollection<LigneContenu>> ligne_fichier in this.liste_lignes_fichier)
                    {
                        Console.WriteLine("*******************Key**********************");
                        Console.WriteLine("Identificateur de la ligne est :" + ligne_fichier.Key);
                        Console.WriteLine("*******************Value**********************");
                        foreach (LigneContenu contenu in ligne_fichier.Value)
                        {
                            contenu.affiche();
                        }
                        Console.WriteLine("-----------------------------------------");
                        Console.WriteLine();
                    }
        
        
        }
    }
}

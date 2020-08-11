using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Xml.Linq;

namespace bamEpplus
{
    public sealed class Parametrage
    {      
            /*Chemins parametres */
            public  String chemin_app = "";
            public  String chemin_shema = "";
            public  String chemin_data = "";
            public  String chemin_genration_excel = "";
            public  String chemin_archive_data = "";
            public  String chemin_template = "";
            public  String chemin_archive_excel = "";
            /*Email Parametres*/
            public bool module_email { get; set; }
            public String adresse { get; set; }
            public String login { get; set; }
            public String password { get; set; }
            public String Host { get; set; }
            public int Port { get; set; }
            public bool EnableSsl { get; set; }
            public bool UseDefaultCredentials { get; set; }
            public String FROM { get; set; }
            public String TO { get; set; }
            public String Subject { get; set; }
            public String Body { get; set; }
            public ICollection<String> CC { get; set; }
            public String userState { get; set; }
            /*Excel parametrages*/
            public String prefixe { get; set; }
            public String nom_fichier_template { get; set; }
            public String code_produit_SMI { get; set; } 
            /* Instanciation de Singleton */
            private static Parametrage parametre = new Parametrage();

       
            private Parametrage()
            {
                initialisation();
            }
            public static Parametrage getInstance()
            {

                return parametre;
            }
            private Parametrage initialisation()
            {
                String chaine_temporaire = "";
                String fichier_config = "";
                chemin_app = ConfigurationManager.AppSettings["CHEMINAPP"];
                fichier_config = chemin_app + @"\CONFG.xml";
                /* par default module mail non actif */
                this.module_email = false; 
                XElement xelement = XElement.Load(@fichier_config);
                var repertoires = from repertoir in xelement.Elements("REPERTOIRES").Elements()

                                    select repertoir;
                Console.WriteLine("Lecture des repertoires .....");
                //Console.WriteLine("---------------------------------");
                foreach (var repertoire in repertoires)
                {
                    chaine_temporaire = "";
                    chaine_temporaire = (string)repertoire.Element("NAME");
                   // chaine_temporaire = chemin_app + @"\" + chaine_temporaire;
                   // Console.WriteLine("ID :" + (string)repertoire.Attribute("ID") + "\t Valeur: " + chaine_temporaire);

                    switch ((string)repertoire.Attribute("ID"))
                    {

                        case "SHEMA":
                            chemin_shema = chaine_temporaire;
                            break;
                        case "EXCELGENERATION":
                            chemin_genration_excel = chaine_temporaire;
                            break;
                        case "DATA":
                            chemin_data = chaine_temporaire;
                            break;
                        case "ARCHIVEDATA":
                            chemin_archive_data = chaine_temporaire;
                            break;
                        case "TEMPLATE":
                            chemin_template = chaine_temporaire;
                            break;
                        case "ARCHIVEEXCEL":
                            chemin_archive_excel = chaine_temporaire;
                            break;
                        default: break;
                    }
                }

                /*Lecture des parametres de l'email */
                /*Reservation de memoire pour CC*/
                CC = new List<String>();
                var proprietes = from propriete in xelement.Elements("PROPRITES").Elements()

                                 select propriete;
                Console.WriteLine("Lecture des parametres .....");
                //  Console.WriteLine("---------------------------------");
                foreach (var propriete in proprietes)
                {
                    chaine_temporaire = "";
                    chaine_temporaire = (string)propriete;
                  //  Console.WriteLine("ID :" + (string) propriete.Attribute("ID") + "\t Valeur: " + chaine_temporaire);
                    switch ((string)propriete.Attribute("ID"))
                    {
                        case "FICHIERTEMPLATE": this.nom_fichier_template = chaine_temporaire;
                            break;
                        case "PREFIXE": this.prefixe = chaine_temporaire;
                            break;
                        case "MODULEEMAIL": if (chaine_temporaire.CompareTo("true") == 0) this.module_email = true;
                            break;
                        case "ADRESSE": this.adresse = chaine_temporaire;
                            break;

                        case "LOGIN": this.login = chaine_temporaire;
                            break;

                        case "PASSWORD": this.password = chaine_temporaire;
                            break;

                        case "HOST": this.Host = chaine_temporaire;
                            break;

                        case "PORT": this.Port = Convert.ToInt32(chaine_temporaire);
                            break;

                        case "ENABLESSL": if (chaine_temporaire.CompareTo("true") == 0) this.EnableSsl = true;
                                          else this.EnableSsl = false;
                            break;

                        case "FROM": this.FROM = chaine_temporaire;
                            break;

                        case "TO": this.TO = chaine_temporaire;
                            break;

                        case "SUBJECT": this.Subject = chaine_temporaire;
                            break;

                        case "BODY": this.Body = chaine_temporaire;
                            break;
                        case "CC": this.CC.Add(chaine_temporaire);
                            break;
                        case "USERSTATE": this.userState = chaine_temporaire;
                            break;
                        case "CODEPRODUITSMI": this.code_produit_SMI = chaine_temporaire;
                            break;
                        case "USEDEFAULTCREDENTIALS": if (chaine_temporaire.CompareTo("true") == 0) this.UseDefaultCredentials = true;
                                                        else this.UseDefaultCredentials = false; 
                            break;
                        default: break;
                    }
                }
                return this;
            }
    }
}

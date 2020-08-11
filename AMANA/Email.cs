/**
 * YON 31/12/2014
 * Class Email : s'occupe de l'envoi de mail 
 * 
 * */


using System;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;
using System.ComponentModel;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace bamEpplus
{
    public sealed class Email
    {   
        private static bool mailSent = false;
        private static bool mailSentSuccess = false;
        private static String fichier_courant = "";
        private static Email email= new Email();
        public static String chemin_trace = "";

        // return une instance de Email 
        public static Email getInstance()
        {
                    return email;
        }
        private static bool traitement_mail(String fichier, Parametrage parametrage)
        {
                SmtpClient client;
                client = new SmtpClient
                {
                    Host = parametrage.Host,
                    Port = parametrage.Port,
                    Timeout = 60000
                };
                Utilitaire.fichier_trace(chemin_trace, " Connection sur l'HOTE : " + client.Host + " Port :" + client.Port);
                client.EnableSsl = parametrage.EnableSsl;
                if (client.EnableSsl)
                {
                         client.Credentials = new System.Net.NetworkCredential(parametrage.login, parametrage.password);
                         Utilitaire.fichier_trace(chemin_trace, " Connection securisé : EnableSsl" + client.EnableSsl);  
                }
 
                client.UseDefaultCredentials = parametrage.UseDefaultCredentials;
                Utilitaire.fichier_trace(chemin_trace, " UseDefaultCredentials : " + client.UseDefaultCredentials + "EnableSsl" + client.EnableSsl);
             
                Utilitaire.fichier_trace(chemin_trace, " FROM " + parametrage.FROM);
                MailAddress from = new MailAddress(parametrage.FROM);
                // Set destinations for the e-mail message.
                Utilitaire.fichier_trace(chemin_trace, " TO " + parametrage.TO);
                MailAddress to = new MailAddress(parametrage.TO);
                // Specify the message content.
                MailMessage message = new MailMessage(from, to);
                foreach (String param in parametrage.CC)
                {
                    message.CC.Add(new MailAddress(param));
                }

                Utilitaire.fichier_trace(chemin_trace, " Body " + parametrage.Body);
                message.Body = parametrage.Body;
                Utilitaire.fichier_trace(chemin_trace, " Subject " + parametrage.Subject);
                message.Subject = parametrage.Subject;

                client.SendCompleted += (s, e) =>
                {
                    SendCompletedCallback(s, e);
                    client.Dispose();
                    message.Dispose();
                };

                ServicePointManager.ServerCertificateValidationCallback = delegate(object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                //string userState = parametrage.userState;
                string userState = parametrage.userState + "-" +Guid.NewGuid().ToString();
                Utilitaire.fichier_trace(chemin_trace, "  UserState" + userState);
                Attachment attachment = new Attachment(@fichier);
                
                 message.Attachments.Add(attachment);
                 Utilitaire.fichier_trace(chemin_trace, " chargement Fichier Excel en piece jointe");
                //Utilitaire.fichier_trace(chemin_trace, " ");
                Utilitaire.fichier_trace(chemin_trace, " Debut de connection SMTP" + parametrage.Subject);
                //client.SendAsync(message, userState);
                try
                {
                    client.SendAsync(message, userState);
                    Utilitaire.fichier_trace(chemin_trace, " Debut d'envoie Asynchrone ..." + parametrage.Subject);
                }
                catch (Exception ex) {
                    Utilitaire.fichier_trace(chemin_trace, "Exception caught in CreateTestMessage2(): {0}" + ex.ToString());
                    client.Dispose();
                    message.Dispose();
                    mailSentSuccess = false;
                    mailSent = true;

                }
               
                
                return true;
        }
        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
                    // Get the unique identifier for this asynchronous operation.
                    String token = (string)e.UserState;
                    Utilitaire.fichier_trace(chemin_trace, " Fin d' envoie Asynchrone resultat est :");

                    if (e.Cancelled)
                    {

                        mailSentSuccess = false;
                        Utilitaire.fichier_trace(chemin_trace, " [{0}] Evoie annulé." + token);
                        //Console.WriteLine("[{0}] Send canceled.", token);
                    }
                    if (e.Error != null)
                    {
                        mailSentSuccess = false;
                        Console.WriteLine(" [{0}] {1}", token, e.Error.ToString());
                        Utilitaire.fichier_trace(chemin_trace, token + e.Error.ToString());
                    }
                    else
                    {
                        mailSentSuccess = true;
                        Utilitaire.fichier_trace(chemin_trace, " Mail a été bien envoyé");
                        Console.WriteLine("Email sent.");
                        
                    }
                    mailSent = true;
        }
        public  bool Envoi_mail(Parametrage parametrage)
        {
            chemin_trace = @parametrage.chemin_archive_data;
            var fichiers = Directory.GetFiles(@parametrage.chemin_genration_excel);
            foreach (String fichier in fichiers)
            {
               // Console.WriteLine(");
                fichier_courant = fichier;
                Utilitaire.fichier_trace(chemin_trace, " Nouveau email : envoie de fichier : " + fichier_courant);
                traitement_mail(fichier, parametrage);
                var mots = fichier_courant.Split('\\');

                String destFile = @parametrage.chemin_archive_excel + @"\" + mots[mots.Length - 1];
                bool resultat = true;
                try
                {
                  while (resultat)
                  {
                  
                        System.Threading.Thread.Sleep(10000);
                       // Utilitaire.fichier_trace(chemin_trace, " Attente envoie mail ...");
                        if (mailSent)
                        {
                            resultat = false;
                            if (mailSentSuccess)
                            {
                                try
                                {
                                    //System.Threading.Thread.Sleep(8000);
                                    System.IO.File.Copy(@fichier_courant, @destFile, true);
                                    // Console.WriteLine("fichier copy dans:." + destFile);
                                    System.Threading.Thread.Sleep(10000);
                                    System.IO.File.Delete(@fichier_courant);
                                    Utilitaire.fichier_trace(chemin_trace, " Fichier EXCEL déplacé vers repertoire d'archive");
                                }
                                catch (Exception ex)
                                {
                                    resultat = false;
                                    mailSentSuccess = false;
                                    //Console.WriteLine("fichier ne peut pas etre supprimé:" + ex.ToString());
                                    Utilitaire.fichier_trace(chemin_trace, " Fichier ne peut pas etre supprimé:" + ex.ToString());

                                }
                            }// fin de test que email a été bien envoyé
                        }//fin de test de fin d'envoie d'email
                   
                     }// fin de while
                }
                catch (Exception ex)
                {
                    resultat = false;
                    mailSent = false;
                    //Console.WriteLine("fichier ne peut pas etre supprimé:" + ex.ToString());
                    Utilitaire.fichier_trace(chemin_trace, "Problème d'envoie mail:" + ex.ToString());

                }
                mailSent = false;
                if (!mailSentSuccess) break;
                mailSentSuccess = false;
            }// fin de foreach 
                   
                return true;
        }// Fin de la methode envoi mail 
    }
}

using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Xml.Linq;
using System.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using bamEpplus;
namespace AMANA
{
    public partial class Service1 : ServiceBase
    {

        private System.Diagnostics.EventLog eventLog1;


        public Service1()
        {
            InitializeComponent();
            OnTimer();
        }
        public void OnDebug()
        {
            OnStart(null);
        }
   
        public void OnTimer()
        {
            //eventLog1.WriteEntry("In OnStart");
            /* Déclaration et Initialisation des parametres*/
            Parametrage parametrage = Parametrage.getInstance();
            Email email = null;
            //ICollection<String> trace = new List<String>();
            // DateTime localDate = DateTime.Now;
            //String chemin_complet_erreur = parametrage.chemin_archive_data + @"\trace\Amana_131_" + localDate.Year + "" + localDate.Month + "" + localDate.Day + "_000001.log";

            Shema_TRA shema = new Shema_TRA();
            Utilitaire.fichier_trace(parametrage.chemin_archive_data, "Initialisation Shema");
            //trace.Add("Initialisation Shema");
            shema.init(parametrage.chemin_shema);
            Utilitaire.fichier_trace(parametrage.chemin_archive_data, "Initialisation Mail");
            //trace.Add("Initialisation Mail");

            if (parametrage.module_email)
            {
                email = Email.getInstance();
            }

            var fichiers = Directory.GetFiles(@parametrage.chemin_data);


            foreach (String fichier in fichiers)
            {
                bool resultat = false;
                var mots = fichier.Split('\\');



                if ((mots[mots.Length - 1].Substring(0, 3).CompareTo("TRA")) == 0)
                {
                    Generation generation = new Generation(shema, "TRA", parametrage);

                    Utilitaire.fichier_trace(parametrage.chemin_archive_data, "Géneration de fichier excel pour fichier : " + fichier);
                    //trace.Add("Géneration de fichier excel pour fichier : " + fichier);
                    resultat = generation.lire_PLAT(fichier);
                    // generation.affiche();
                    if (resultat == true)
                    {
                        resultat = generation.ecrire_excel();
                    }

                    String destFile = @parametrage.chemin_archive_data + @"\" + mots[mots.Length - 1];
                    if (resultat == true)
                    {
                        System.IO.File.Copy(fichier, destFile, true);
                        System.IO.File.Delete(fichier);
                    }

                }

            }// fin de foreach

            if (parametrage.module_email)
            {
                Utilitaire.fichier_trace(parametrage.chemin_archive_data, "**** Module Mail actif *****");
                // trace.Add("Debut envoi mail");
                email.Envoi_mail(parametrage);
                //trace.Add("Fin d 'envoi mail");
                Utilitaire.fichier_trace(parametrage.chemin_archive_data, "**** FIN Traitements des e-mails  ***** ");
            }
            // System.IO.File.AppendAllLines(@chemin_complet_erreur, trace);
        }


        public Service1(string[] args)       {
            InitializeComponent();
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = Double.Parse(((String)ConfigurationManager.AppSettings["TIMESPERIODE"]));//30000; // 60 seconds
            //Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();
            // Set up a timer to trigger every minute.

        }

        protected override void OnStart(string[] args)
        {
            
           // eventLog1.WriteEntry("In OnStart");
            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
            

           
        }

        protected override void OnStop()
        {
            //eventLog1.WriteEntry("In OnStop");
        }

        protected override void OnContinue()
        {
            //eventLog1.WriteEntry("In OnContinue");
        }
        protected override void OnPause()
        {
            //eventLog1.WriteEntry("In OnPause");
           
        }

        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public long dwServiceType;
            public ServiceState dwCurrentState;
            public long dwControlsAccepted;
            public long dwWin32ExitCode;
            public long dwServiceSpecificExitCode;
            public long dwCheckPoint;
            public long dwWaitHint;
        };
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);



    }
}

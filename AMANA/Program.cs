using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace AMANA
{
    static class Program
    {
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        static void Main(string[] args)
        {
#if DEBUG
            Service1 myService = new Service1();
            myService.OnDebug();
            System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);
            
#else
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
			{ 
				new Service1(args)
			};
            ServiceBase.Run(ServicesToRun);
#endif
        }
    } 
}




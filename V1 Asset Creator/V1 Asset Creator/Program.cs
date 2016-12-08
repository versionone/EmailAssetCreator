using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace V1_Asset_Creator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {

#if DEBUG
            V1_Asset_Creator AssetCreatorService = new V1_Asset_Creator();
            AssetCreatorService.DuringDebug();
            System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);

#else
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
            { 
                new V1_Asset_Creator() 
            };
            ServiceBase.Run(ServicesToRun);
            //ServiceBase.Run(new V1_Asset_Creator()/*ServicesToRun*/);

            
           
#endif
        }
    }
}

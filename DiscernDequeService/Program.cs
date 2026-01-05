using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace DiscernDequeService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 
#if DEBUG
        static ConsoleWorker m_Worker = null;
#endif
        static void Main()
        {
#if DEBUG
            m_Worker = new ConsoleWorker();
            m_Worker.Start();
            Console.WriteLine("ConsoleWorker started...");
            Console.WriteLine("Press [CTRL-C] to close console...");
            Console.CancelKeyPress += new ConsoleCancelEventHandler(Console_CancelKeyPress);
#else
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new DiscernDequeService()
            };
            
            ServiceBase.Run(ServicesToRun);
#endif
        }
#if DEBUG
        static void Console_CancelKeyPress(object sender, ConsoleCancelEventArgs e)
        {
            Console.WriteLine("Exiting...");
            //m_Worker.Stop();
        }
#endif
    }
}

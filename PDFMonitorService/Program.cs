using System;
using System.Diagnostics;
using System.ServiceProcess;

namespace PDFMonitorServiceNamespace
{
    static class Program
    {
        static void Main(string[] args)
        {
            try
            {
                PDFMonitorService.WriteToEventLog("Program started.", EventLogEntryType.Information, 1002);

                if (Environment.UserInteractive)
                {
                    // Modo consola
                    PDFMonitorService.WriteToEventLog("Running in console mode.", EventLogEntryType.Information, 1003);
                    RunAsConsole(args);
                }
                else
                {
                    // Modo servicio
                    PDFMonitorService.WriteToEventLog("Running as a service.", EventLogEntryType.Information, 1004);
                    ServiceBase[] ServicesToRun;
                    ServicesToRun = new ServiceBase[]
                    {
                        new PDFMonitorService()
                    };
                    ServiceBase.Run(ServicesToRun);
                }
            }
            catch (Exception ex)
            {
                PDFMonitorService.WriteToEventLog($"Unhandled exception: {ex.Message}", EventLogEntryType.Error, 7008);
            }
        }

        private static void RunAsConsole(string[] args)
        {
            try
            {
                PDFMonitorService.WriteToEventLog("Initializing service...", EventLogEntryType.Information, 1005);

                PDFMonitorService service = new PDFMonitorService();
                service.StartService(args);

                PDFMonitorService.WriteToEventLog("Service started. Running in console mode.", EventLogEntryType.Information, 2002);

                // Mantener el servicio en ejecución hasta que se cierre manualmente
                System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);
            }
            catch (Exception ex)
            {
                PDFMonitorService.WriteToEventLog($"Error running as console: {ex.Message}", EventLogEntryType.Error, 7009);
            }
        }
    }
}

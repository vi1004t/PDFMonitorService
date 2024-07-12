using System;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Timers;
using System.Configuration;
using System.Linq;
using System.Xml.Linq;
using PdfiumViewer;
using System.Drawing.Printing;


namespace PDFMonitorServiceNamespace
{
    public partial class PDFMonitorService : ServiceBase
    {
        private Timer timer;
        private string monitorPath;
        private int intervalInSeconds;
        private string fileSuffix;
        private string printerName;
        private int retryDelay;
        private int printTimeout;
        private bool enablePrintRetry;
        private int maxPrintRetries;
        private int maxTimerRetries;
        private static string EventSource;
        private bool debugMode;

        private const string EventLogName = "PDFMonitorServiceLog";
        private const string DefaultEventSourceName = "PDFMonitorServiceSource1";
        private const string DefaultMonitorPath = @"C:\Path\To\Monitor";
        private const int DefaultIntervalInSeconds = 3600;
        private const string DefaultFileSuffix = "_processed";
        private const string DefaultPrinterName = "Microsoft Print to PDF";
        private const int DefaultRetryDelayInMilliseconds = 1000;
        private const int DefaultPrintTimeoutInMilliseconds = 10000;
        private const bool DefaultEnablePrintRetry = true;
        private const int DefaultMaxPrintRetries = 3;
        private const int DefaultMaxTimerRetries = 0;
        private const bool DefaultdebugMode = false;

        public PDFMonitorService()
        {
            InitializeComponent();
            try
            {
                EnsureConfigFile();
                ReadConfiguration();
            }
            catch (Exception ex)
            {
                WriteToEventLog($"Error during initialization: {ex.Message}", EventLogEntryType.Error, 7001);
                throw;
            }
        }

        private void EnsureConfigFile()
        {
            bool configUpdated = false;
            string configFilePath = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;
            XDocument config;

            if (!File.Exists(configFilePath))
            {
                config = new XDocument(new XElement("configuration",
                    new XElement("appSettings",
                    
                        new XElement("add", new XAttribute("key", "EventSourceName"), new XAttribute("value", DefaultEventSourceName)),
                        new XElement("add", new XAttribute("key", "MonitorPath"), new XAttribute("value", DefaultMonitorPath)),
                        new XElement("add", new XAttribute("key", "IntervalInSeconds"), new XAttribute("value", DefaultIntervalInSeconds)),
                        new XElement("add", new XAttribute("key", "FileSuffix"), new XAttribute("value", DefaultFileSuffix)),
                        new XElement("add", new XAttribute("key", "PrinterName"), new XAttribute("value", DefaultPrinterName)),
                        new XElement("add", new XAttribute("key", "RetryDelayInMilliseconds"), new XAttribute("value", DefaultRetryDelayInMilliseconds)),
                        new XElement("add", new XAttribute("key", "PrintTimeoutInMilliseconds"), new XAttribute("value", DefaultPrintTimeoutInMilliseconds)),
                        new XElement("add", new XAttribute("key", "EnablePrintRetry"), new XAttribute("value", DefaultEnablePrintRetry)),
                        new XElement("add", new XAttribute("key", "MaxPrintRetries"), new XAttribute("value", DefaultMaxPrintRetries)),
                        new XElement("add", new XAttribute("key", "MaxTimerRetries"), new XAttribute("value", DefaultMaxTimerRetries)),
                        new XElement("add", new XAttribute("key", "DebugMode"), new XAttribute("value", DefaultdebugMode))
                    )
                ));
                configUpdated = true;
            }
            else
            {
                config = XDocument.Load(configFilePath);
                var appSettings = config.Element("configuration").Element("appSettings");
                if (appSettings != null)
                {
                    configUpdated |= EnsureConfigElement(appSettings, "EventSourceName", DefaultEventSourceName);
                    configUpdated |= EnsureConfigElement(appSettings, "MonitorPath", DefaultMonitorPath);
                    configUpdated |= EnsureConfigElement(appSettings, "IntervalInSeconds", DefaultIntervalInSeconds.ToString());
                    configUpdated |= EnsureConfigElement(appSettings, "FileSuffix", DefaultFileSuffix);
                    configUpdated |= EnsureConfigElement(appSettings, "PrinterName", DefaultPrinterName);
                    configUpdated |= EnsureConfigElement(appSettings, "RetryDelayInMilliseconds", DefaultRetryDelayInMilliseconds.ToString());
                    configUpdated |= EnsureConfigElement(appSettings, "PrintTimeoutInMilliseconds", DefaultPrintTimeoutInMilliseconds.ToString());
                    configUpdated |= EnsureConfigElement(appSettings, "EnablePrintRetry", DefaultEnablePrintRetry.ToString());
                    configUpdated |= EnsureConfigElement(appSettings, "MaxPrintRetries", DefaultMaxPrintRetries.ToString());
                    configUpdated |= EnsureConfigElement(appSettings, "MaxTimerRetries", DefaultMaxTimerRetries.ToString());
                    configUpdated |= EnsureConfigElement(appSettings, "DebugMode", DefaultdebugMode.ToString());

                }
            }

            if (configUpdated)
            {
                config.Save(configFilePath);
                WriteToEventLog("Configuration file created/updated with default values. Please review and restart the service.", EventLogEntryType.Information, 1001);
                Stop();
            }
        }

        private bool EnsureConfigElement(XElement appSettings, string key, string defaultValue)
        {
            var element = appSettings.Elements("add").FirstOrDefault(e => e.Attribute("key")?.Value == key);
            if (element == null)
            {
                appSettings.Add(new XElement("add", new XAttribute("key", key), new XAttribute("value", defaultValue)));
                return true;
            }
            return false;
        }

        private void ReadConfiguration()
        {

            EventSource = ConfigurationManager.AppSettings["EventSourceName"];
            monitorPath = ConfigurationManager.AppSettings["MonitorPath"];
            intervalInSeconds = int.Parse(ConfigurationManager.AppSettings["IntervalInSeconds"]);
            fileSuffix = ConfigurationManager.AppSettings["FileSuffix"];
            printerName = ConfigurationManager.AppSettings["PrinterName"];
            retryDelay = int.Parse(ConfigurationManager.AppSettings["RetryDelayInMilliseconds"]);
            printTimeout = int.Parse(ConfigurationManager.AppSettings["PrintTimeoutInMilliseconds"]);
            enablePrintRetry = bool.Parse(ConfigurationManager.AppSettings["EnablePrintRetry"]);
            maxPrintRetries = int.Parse(ConfigurationManager.AppSettings["MaxPrintRetries"]);
            maxTimerRetries = int.Parse(ConfigurationManager.AppSettings["MaxTimerRetries"]);
            debugMode = bool.Parse(ConfigurationManager.AppSettings["DebugMode"]);
        }

        protected override void OnStart(string[] args)
        {
            WriteToEventLog("Service started.", EventLogEntryType.Information, 2000);
            InitializeTimer();
        }

        private void InitializeTimer()
        {
            int retryCount = 0;

            while (maxTimerRetries == 0 || retryCount < maxTimerRetries)
            {
                try
                {
                    timer = new Timer();
                    timer.Interval = intervalInSeconds * 1000;
                    timer.Elapsed += OnElapsedTime;
                    timer.AutoReset = false; // Desactivar el reinicio automático
                    timer.Start();
                    if (debugMode)
                    {
                        WriteToEventLog("Timer initialized successfully.", EventLogEntryType.Information, 3000);
                    }
                    return;
                }
                catch (Exception ex)
                {
                    WriteToEventLog($"Error initializing timer: {ex.Message}", EventLogEntryType.Error, 7002);
                    retryCount++;
                    System.Threading.Thread.Sleep(retryDelay); // Esperar antes de reintentar
                }
            }

            WriteToEventLog("Failed to initialize timer after max retries.", EventLogEntryType.Error, 7003);
            Stop();
        }

        private void OnElapsedTime(object sender, ElapsedEventArgs e)
        {
            try
            {
                ProcessFiles();
            }
            catch (Exception ex)
            {
                WriteToEventLog($"Error processing files: {ex.Message}", EventLogEntryType.Error, 7004);
            }
            finally
            {
                RestartTimer();
            }
        }

        private void RestartTimer()
        {
            InitializeTimer(); // Llamamos directamente a InitializeTimer para manejar la lógica de reintentos
        }

        private void ProcessFiles()
        {
            try
            {
                var files = Directory.GetFiles(monitorPath, "*.pdf")
                                     .Where(file => !file.EndsWith(fileSuffix + ".pdf"));

                foreach (var file in files)
                {
                    if (debugMode)
                    {
                        WriteToEventLog("Start print document...", EventLogEntryType.Information, 3001);
                    }

                    PrintFile(file, printerName);

                    if (debugMode)
                    {
                        WriteToEventLog("Start rename document...", EventLogEntryType.Information, 3002);
                    }

                    RenameFile(file);
                }
            }
            catch (Exception ex)
            {
                WriteToEventLog($"Error processing files: {ex.Message}", EventLogEntryType.Error, 7005);
                throw; // Volver a lanzar la excepción para capturarla en OnElapsedTime
            }
        }

        //función para imprimir ficheros PDF que necesita tener instalado un visor PDF como Adobe Acrobat Reades u otros
        /*private void PrintFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                WriteToEventLog($"File not found: {filePath}", EventLogEntryType.Error, 5001);
                return;
            }

            int currentRetry = 0;

            while (!enablePrintRetry || currentRetry < maxPrintRetries)
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = filePath,
                        Verb = "printto",
                        Arguments = "\"" + printerName + "\"",
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    };
                    Process p = Process.Start(psi);
                    p.WaitForExit(printTimeout);
                    if (!p.HasExited)
                    {
                        p.Kill();
                        throw new Exception("Timeout while printing.");
                    }
                    WriteToEventLog($"Printed file: {filePath}", EventLogEntryType.Information, 6000);
                    return; // Salir si la impresión fue exitosa
                }
                catch (Exception ex)
                {
                    currentRetry++;
                    WriteToEventLog($"Error printing file (attempt {currentRetry}): {ex.Message}", EventLogEntryType.Error, 5002);
                    if (!enablePrintRetry || currentRetry >= maxPrintRetries)
                    {
                        WriteToEventLog($"Failed to print file after {currentRetry} attempts: {filePath}", EventLogEntryType.Error, 5003);
                        throw; // Volver a lanzar la excepción si se agotaron los reintentos
                    }
                    System.Threading.Thread.Sleep(retryDelay); // Esperar antes de reintentar
                }
            }
        }*/

        private void PrintFile(string filePath, string printerName)
        {
            if (!File.Exists(filePath))
            {
                WriteToEventLog($"File not found: {filePath}", EventLogEntryType.Error, 5001);
                return;
            }

            int currentRetry = 0;

            while (!enablePrintRetry || currentRetry < maxPrintRetries)
            {
                try
                {
                    using (var document = PdfDocument.Load(filePath))
                    {
                        using (var printDocument = document.CreatePrintDocument())
                        {
                            printDocument.PrinterSettings.PrinterName = printerName;
                            printDocument.PrintController = new StandardPrintController(); // No mostrar la interfaz de impresión

                            printDocument.Print();
                        }
                    }

                    WriteToEventLog($"Printed file: {filePath} to printer: {printerName}", EventLogEntryType.Information, 6000);
                    return; // Salir si la impresión fue exitosa
                }
                catch (Exception ex)
                {
                    currentRetry++;
                    WriteToEventLog($"Error printing file (attempt {currentRetry}): {ex.Message}", EventLogEntryType.Error, 5002);
                    if (!enablePrintRetry || currentRetry >= maxPrintRetries)
                    {
                        WriteToEventLog($"Failed to print file after {currentRetry} attempts: {filePath}", EventLogEntryType.Error, 5003);
                        throw; // Volver a lanzar la excepción si se agotaron los reintentos
                    }
                    System.Threading.Thread.Sleep(retryDelay); // Esperar antes de reintentar
                }
            }
        }


        private void RenameFile(string filePath)
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                var newFileName = Path.GetFileNameWithoutExtension(filePath) + "_" + timestamp + fileSuffix + ".pdf";
                var newFilePath = Path.Combine(monitorPath, newFileName);
                File.Move(filePath, newFilePath);
                WriteToEventLog($"Renamed file to: {newFilePath}", EventLogEntryType.Information, 6000);
            }
            catch (Exception ex)
            {
                WriteToEventLog($"Error renaming file: {ex.Message}", EventLogEntryType.Error, 7006);
                throw; // Volver a lanzar la excepción para capturarla en ProcessFiles
            }
        }

        protected override void OnStop()
        {
            try
            {
                timer?.Stop();
                WriteToEventLog("Service stopped.", EventLogEntryType.Information, 2001);
            }
            catch (Exception ex)
            {
                WriteToEventLog($"Error stopping service: {ex.Message}", EventLogEntryType.Error, 7007);
            }
        }

        public static void WriteToEventLog(string message, EventLogEntryType type = EventLogEntryType.Information, int eventID = 0)
        {


            try
            {
                // Crear el origen del evento si no existe
                if (!EventLog.SourceExists(EventSource))
                {
                    EventLog.CreateEventSource(EventSource, EventLogName);
                }

                // Escribir en el registro de eventos
                EventLog.WriteEntry(EventSource, message, type, eventID);

            }
            catch (Exception ex)
            {
                // Intentar escribir el error en el registro de eventos usando un origen genérico
                try
                {
                    EventLog.WriteEntry("Application", $"Error writing to event log: {ex.Message}", EventLogEntryType.Error, eventID);
                }
                catch
                {
                    // Si falla, no se puede hacer mucho más
                }
            }
        }

        // Métodos públicos para iniciar y detener el servicio en modo consola
        public void StartService(string[] args)
        {
            OnStart(args);
        }

        public void StopService()
        {
            OnStop();
        }
    }
}
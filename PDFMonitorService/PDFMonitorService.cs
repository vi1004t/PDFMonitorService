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
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Threading.Tasks;


namespace PDFMonitorServiceNamespace
{
    public partial class PDFMonitorService : ServiceBase
    {
        private Dictionary<string, TimerInfo> timers = new Dictionary<string, TimerInfo>(); //Diccionario para tratar todos los timers (uno por cada carpeta)
        private Dictionary<string, string> folderPrinterMappings = new Dictionary<string, string>(); // Diccionario para listar todas las duplas carpeta-impresora
        private Dictionary<string, HashSet<string>> printQueues = new Dictionary<string, HashSet<string>>(); // Diccionario para gestionar las colas de impresión de cada timer

        private int intervalInSeconds; //intervalo de ejecución de los timers
        private string fileSuffix; //sufijo para el renombre de documentos tratados
        private int retryDelay; //retraso en milisegundos para volver a intentar lanzar el timer
        private int printTimeout; //restraso en segundos para dar la impresión KO
        private bool enablePrintRetry; //habilitar reimpresión en caso de fallo
        private bool moveToErrorFolder; //habilitar mover pdf's corruptos a una carpeta
        private int maxPrintRetries; //máximo de intentos de impresión
        private int maxTimerRetries; //máximo de intentos de lanzar un timer
        private static string EventSource = "PDFMonitorServiceSource"; //temporal hasta que se defina el correcto. Para evitar mensajes de errores en el log de windows
        private bool debugMode; //modo debug
        private Timer supervisionTimer; //Monitor de supervisión de los timers
        private int supervisionIntervalInSeconds; // Intervalo de supervisión
        private int selectTimeInterval; //Filtro de tiempo en el listado de ficheros

        //Valores por defecto
        private const string EventLogName = "PDFMonitorServiceLog"; //nombre de los eventos
        private const string DefaultEventSourceName = "PDFMonitorServiceSource"; //nombre del origen de los eventos
        private const string DefaultMonitorPath = @"C:\Path\To\Monitor";
        private const int DefaultIntervalInSeconds = 60;
        private const string DefaultFileSuffix = "_processed";
        private const string DefaultPrinterName = "Microsoft Print to PDF";
        private const int DefaultRetryDelayInMilliseconds = 1000;
        private const int DefaultPrintTimeoutInMilliseconds = 10000;
        private const bool DefaultEnablePrintRetry = true;
        private const bool DefaultMoveToErrorFolder = false;
        private const int DefaultMaxPrintRetries = 3;
        private const int DefaultMaxTimerRetries = 0;
        private const bool DefaultdebugMode = false;
        private const int DefaultSupervisionIntervalInSeconds = 120; // Intervalo de supervisión
        private const int DefaultSelectTimeInterval = 60;


        /* Función de inicialización */
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
        /*Función que comprueba que el fichero de configuración esté OK, sino existe lo crea 
         * y si hay algún dato que no existe entonces coge su valor por defecto*/
        private void EnsureConfigFile()
        {
            bool configUpdated = false;
            string configFilePath = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;
            XDocument config;

            if (!File.Exists(configFilePath))
            {
                config = new XDocument(new XElement("configuration",
                    new XElement("configSections",
                            new XElement("section", new XAttribute("name", "folderPrinterMappings"), new XAttribute("type", "System.Configuration.NameValueSectionHandler"))
                        ),
                        new XElement("folderPrinterMappings",
                            new XElement("add", new XAttribute("key", DefaultMonitorPath), new XAttribute("value", DefaultPrinterName))
                        ),
                    new XElement("appSettings",

                        new XElement("add", new XAttribute("key", "EventSourceName"), new XAttribute("value", DefaultEventSourceName)),
                        new XElement("add", new XAttribute("key", "IntervalInSeconds"), new XAttribute("value", DefaultIntervalInSeconds)),
                        new XElement("add", new XAttribute("key", "FileSuffix"), new XAttribute("value", DefaultFileSuffix)),
                        new XElement("add", new XAttribute("key", "RetryDelayInMilliseconds"), new XAttribute("value", DefaultRetryDelayInMilliseconds)),
                        new XElement("add", new XAttribute("key", "PrintTimeoutInMilliseconds"), new XAttribute("value", DefaultPrintTimeoutInMilliseconds)),
                        new XElement("add", new XAttribute("key", "EnablePrintRetry"), new XAttribute("value", DefaultEnablePrintRetry)),
                        new XElement("add", new XAttribute("key", "MaxPrintRetries"), new XAttribute("value", DefaultMaxPrintRetries)),
                        new XElement("add", new XAttribute("key", "MaxTimerRetries"), new XAttribute("value", DefaultMaxTimerRetries)),
                        new XElement("add", new XAttribute("key", "DebugMode"), new XAttribute("value", DefaultdebugMode)),
                        new XElement("add", new XAttribute("key", "SupervisionIntervalInSeconds"), new XAttribute("value", DefaultSupervisionIntervalInSeconds)),
                        new XElement("add", new XAttribute("key", "SelectTimeInterval"), new XAttribute("value", DefaultSelectTimeInterval)),
                        new XElement("add", new XAttribute("key", "MoveToErrorFolder"), new XAttribute("value", DefaultMoveToErrorFolder))

                    )
                ));
                configUpdated = true;
            }
            else
            {
                config = XDocument.Load(configFilePath);

                // Verificar y crear configSections si no existe
                configUpdated |= EnsureConfigSection(config, "configSections");
                var configSections = config.Element("configuration").Element("configSections");

                // Verificar y crear la sección folderPrinterMappings si no existe
                var folderPrinterMappingSection = configSections.Elements("section").FirstOrDefault(e => e.Attribute("name")?.Value == "folderPrinterMappings");
                if (folderPrinterMappingSection == null)
                {
                    configSections.Add(new XElement("section", new XAttribute("name", "folderPrinterMappings"), new XAttribute("type", "System.Configuration.NameValueSectionHandler")));
                    configUpdated = true;
                }

                // Verificar y crear folderPrinterMappings si no existe
                configUpdated |= EnsureConfigSection(config, "folderPrinterMappings");
                var folderPrinterMappings = config.Element("configuration").Element("folderPrinterMappings");
                //comentado porque lo que hace es generar un registro no valido ya que el valor de DefaultMonitorPath es un ejemplo
                //y al ejecutar el monitor da siempre error
                //configUpdated |= EnsureConfigElement(folderPrinterMappings, DefaultMonitorPath, DefaultPrinterName);

                // Verificar y crear appSettings si no existe
                configUpdated |= EnsureConfigSection(config, "appSettings");
                var appSettings = config.Element("configuration").Element("appSettings");
                configUpdated |= EnsureConfigElement(appSettings, "EventSourceName", DefaultEventSourceName);
                configUpdated |= EnsureConfigElement(appSettings, "IntervalInSeconds", DefaultIntervalInSeconds.ToString());
                configUpdated |= EnsureConfigElement(appSettings, "FileSuffix", DefaultFileSuffix);
                configUpdated |= EnsureConfigElement(appSettings, "RetryDelayInMilliseconds", DefaultRetryDelayInMilliseconds.ToString());
                configUpdated |= EnsureConfigElement(appSettings, "PrintTimeoutInMilliseconds", DefaultPrintTimeoutInMilliseconds.ToString());
                configUpdated |= EnsureConfigElement(appSettings, "EnablePrintRetry", DefaultEnablePrintRetry.ToString());
                configUpdated |= EnsureConfigElement(appSettings, "MaxPrintRetries", DefaultMaxPrintRetries.ToString());
                configUpdated |= EnsureConfigElement(appSettings, "MaxTimerRetries", DefaultMaxTimerRetries.ToString());
                configUpdated |= EnsureConfigElement(appSettings, "DebugMode", DefaultdebugMode.ToString());
                configUpdated |= EnsureConfigElement(appSettings, "SupervisionIntervalInSeconds", DefaultSupervisionIntervalInSeconds.ToString());
                configUpdated |= EnsureConfigElement(appSettings, "SelectTimeInterval", DefaultSelectTimeInterval.ToString());
                configUpdated |= EnsureConfigElement(appSettings, "MoveToErrorFolder", DefaultMoveToErrorFolder.ToString());
            }

            if (configUpdated)
            {
                config.Save(configFilePath);
                WriteToEventLog("Configuration file created/updated with default values. Please review and restart the service.", EventLogEntryType.Information, 1001);
                Stop();
            }
        }
        /*Función auxiliar para crear elementos de configuración*/
        private bool EnsureConfigElement(XElement settings, string key, string defaultValue)
        {
            var element = settings.Elements("add").FirstOrDefault(e => e.Attribute("key")?.Value == key);
            if (element == null)
            {
                settings.Add(new XElement("add", new XAttribute("key", key), new XAttribute("value", defaultValue)));
                return true;
            }
            return false;
        }
        /*Función auxiliar para crear elementos de configuración*/
        private bool EnsureConfigSection(XDocument config, string sectionName)
        {
            var section = config.Element("configuration").Element(sectionName);
            if (section == null)
            {
                config.Root.Add(new XElement(sectionName));
                return true;
            }
            return false;
        }
        /*Función para leer las variables de configuración*/
        private void ReadConfiguration()
        {
            EventSource = ConfigurationManager.AppSettings["EventSourceName"];
            intervalInSeconds = int.Parse(ConfigurationManager.AppSettings["IntervalInSeconds"]);
            fileSuffix = ConfigurationManager.AppSettings["FileSuffix"];
            retryDelay = int.Parse(ConfigurationManager.AppSettings["RetryDelayInMilliseconds"]);
            printTimeout = int.Parse(ConfigurationManager.AppSettings["PrintTimeoutInMilliseconds"]);
            enablePrintRetry = bool.Parse(ConfigurationManager.AppSettings["EnablePrintRetry"]);
            maxPrintRetries = int.Parse(ConfigurationManager.AppSettings["MaxPrintRetries"]);
            maxTimerRetries = int.Parse(ConfigurationManager.AppSettings["MaxTimerRetries"]);
            debugMode = bool.Parse(ConfigurationManager.AppSettings["DebugMode"]);
            supervisionIntervalInSeconds = int.Parse(ConfigurationManager.AppSettings["SupervisionIntervalInSeconds"]);
            selectTimeInterval = int.Parse(ConfigurationManager.AppSettings["SelectTimeInterval"]);
            moveToErrorFolder= bool.Parse(ConfigurationManager.AppSettings["MoveToErrorFolder"]);

            var folderPrinterMappingsSection = ConfigurationManager.GetSection("folderPrinterMappings") as NameValueCollection;
            if (folderPrinterMappingsSection != null)
            {
                foreach (string folder in folderPrinterMappingsSection.AllKeys)
                {
                    string printer = folderPrinterMappingsSection[folder];
                    folderPrinterMappings.Add(folder, printer);
                    WriteToEventLog($"Loaded mapping: Folder={folder}, Printer={printer}", EventLogEntryType.Information, 3004);
                }
            }
        }
        /*Función que se ejecuta al iniciarse, después del constructor */
        protected override void OnStart(string[] args)
        {
            WriteToEventLog("Service started.", EventLogEntryType.Information, 2000);
            foreach (var mapping in folderPrinterMappings)
            {
                // Inicializa un temporizador por cada mapeo de carpeta-impresora
                WriteToEventLog($"Starting {mapping.Key} on {mapping.Value}.", EventLogEntryType.Information, 2000);
                Task.Run(() => InitializeTimer(mapping.Key, mapping.Value));
            }
            // Iniciar el temporizador de supervisión
            StartSupervisionTimer();
        }
        /*Función para inicializar un timer, recibe como parámetro la carpeta a monitorizar y la impresora por la que imprimir*/
        private void InitializeTimer(string folder, string printer)
        {
            int retryCount = 0;

            while (maxTimerRetries == 0 || retryCount < maxTimerRetries)
            {
                try
                {
                    Timer timer = new Timer();
                    timer.Interval = intervalInSeconds * 1000;
                    timer.Elapsed += (sender, e) => OnElapsedTime(sender, e, folder, printer);
                    timer.AutoReset = false; // Desactivar el reinicio automático
                    timer.Start();
                    lock (timers)
                    {
                        timers[folder] = new TimerInfo { Timer = timer, LastRunTime = DateTime.Now, Folder = folder, Printer = printer };
                    }

                    if (debugMode)
                    {
                        WriteToEventLog($"Timer initialized successfully for folder {folder}.", EventLogEntryType.Information, 3000);
                    }
                    return;
                }
                catch (Exception ex)
                {
                    WriteToEventLog($"Error initializing timer for folder {folder}: {ex.Message}", EventLogEntryType.Error, 7002);
                    retryCount++;
                    System.Threading.Thread.Sleep(retryDelay); // Esperar antes de reintentar
                }
            }

            WriteToEventLog($"Failed to initialize timer for folder {folder} after max retries.", EventLogEntryType.Error, 7003);
            Stop();
        }
        /*Función que se ejecuta al terminarse el timer*/
        private void OnElapsedTime(object sender, ElapsedEventArgs e, string folder, string printer)
        {
            try
            {
                _ = ProcessFilesAsync(folder, printer); // Llama a ProcessFilesAsync y olvídate de la tarea.
                lock (timers)
                {
                    if (timers.ContainsKey(folder))
                    {
                        timers[folder].LastRunTime = DateTime.Now;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToEventLog($"Error processing files in folder {folder}: {ex.Message}", EventLogEntryType.Error, 7004);
            }
            finally
            {
                RestartTimer(folder, printer);
            }
        }
        /*Función que llama de nuevo al timer*/
        private void RestartTimer(string folder, string printer)
        {
            Task.Run(() => InitializeTimer(folder, printer)); // Llamamos directamente a InitializeTimer para manejar la lógica de reintentos
        }
        /*Función para inicializar el timer de supervisión*/
        private void StartSupervisionTimer()
        {
            supervisionTimer = new Timer();
            supervisionTimer.Interval = supervisionIntervalInSeconds * 1000; // Intervalo de supervisión
            supervisionTimer.Elapsed += SupervisionTimerElapsed;
            supervisionTimer.AutoReset = true;
            supervisionTimer.Start();
        }
        /*Función que se ejecuta al terminar el timer de supervisión*/
        private void SupervisionTimerElapsed(object sender, ElapsedEventArgs e)
        {
            DateTime now = DateTime.Now;
            lock (timers)
            {
                foreach (var timerInfo in timers.Values)
                {
                    if ((now - timerInfo.LastRunTime).TotalSeconds > intervalInSeconds * 2)
                    {
                        WriteToEventLog($"Restarting timer for folder {timerInfo.Folder} due to inactivity.", EventLogEntryType.Warning, 7006);
                        RestartTimer(timerInfo.Folder, timerInfo.Printer);
                    }
                }
            }

            // Supervisar el temporizador de supervisión
            if ((now - e.SignalTime).TotalSeconds > supervisionIntervalInSeconds * 2)
            {
                WriteToEventLog("Restarting supervision timer due to inactivity.", EventLogEntryType.Warning, 7007);
                StartSupervisionTimer();
            }
        }
        /*Función que se ejecuta en el OnElapsedTime, responsable de generar las llamadas a ProcesFile por cada fichero*/
        private async Task ProcessFilesAsync(string folder, string printer)
        {
            try
            {
                if (!printQueues.ContainsKey(folder))
                {
                    printQueues[folder] = new HashSet<string>();
                }

                var files = Directory.GetFiles(folder, "*.pdf")
                                     .Where(file => !file.EndsWith(fileSuffix + ".pdf"))
                                     .Where(file => new FileInfo(file).CreationTime >= DateTime.Now.AddMinutes(-selectTimeInterval) ||
                                                    new FileInfo(file).LastWriteTime >= DateTime.Now.AddMinutes(-selectTimeInterval))
                                     .Where(file => !printQueues[folder].Contains(file)) // Filtrar archivos que ya están en la cola
                                     .ToList();

                foreach (var file in files)
                {
                    lock (printQueues)
                    {
                        printQueues[folder].Add(file);
                    }
                    WriteToEventLog($"Queued file for processing: {file}", EventLogEntryType.Information, 3006);
                }

                var tasks = files.Select(file => Task.Run(() => ProcessFile(file, printer, folder))).ToList();
                await Task.WhenAll(tasks);
            }
            catch (Exception ex)
            {
                WriteToEventLog($"Error processing files in folder {folder}: {ex.Message}", EventLogEntryType.Error, 7005);
            }
        }
        /*Función que orquesta la impresión y renombre de un fichero*/
        private void ProcessFile(string file, string printer, string folder)
        {
            try
            {
                if (debugMode)
                {
                    WriteToEventLog("Start print document...", EventLogEntryType.Information, 3001);
                }

                PrintFile(file, printer);

                if (debugMode)
                {
                    WriteToEventLog("Start rename document...", EventLogEntryType.Information, 3002);
                }

                RenameFile(file);
            }
            catch (Exception ex)
            {
                WriteToEventLog($"Error processing file {file}: {ex.Message}", EventLogEntryType.Error, 5002);
                if (moveToErrorFolder)
                {
                    HandleErrorFile(file, ex);
                }
                
            }
            finally
            {
                lock (printQueues)
                {
                    if (printQueues.ContainsKey(folder))
                    {
                        printQueues[folder].Remove(file);
                        WriteToEventLog($"Removed file from queue: {file}", EventLogEntryType.Information, 3007);
                    }
                }
            }
        }
        /*Función auxiliar para mover de carpeta los pdf's corruptos*/
        private void HandleErrorFile(string file, Exception ex)
        {
            try
            {
                string errorFolder = Path.Combine(Path.GetDirectoryName(file), "ErrorFiles");
                if (!Directory.Exists(errorFolder))
                {
                    Directory.CreateDirectory(errorFolder);
                }

                string errorFilePath = Path.Combine(errorFolder, Path.GetFileName(file));
                File.Move(file, errorFilePath);

                WriteToEventLog($"Moved problematic file to {errorFilePath} due to error: {ex.Message}", EventLogEntryType.Warning, 5003);
            }
            catch (Exception moveEx)
            {
                WriteToEventLog($"Failed to move problematic file {file} due to error: {moveEx.Message}", EventLogEntryType.Error, 5004);
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
        /*Función de impresión de PDF, hace uso de la librería pdfiumviewer*/
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
        /*Función para renombrar ficheros*/
        private void RenameFile(string filePath)
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                var directory = Path.GetDirectoryName(filePath);
                var newFileName = Path.GetFileNameWithoutExtension(filePath) + "_" + timestamp + fileSuffix + ".pdf";
                var newFilePath = Path.Combine(directory, newFileName);
                File.Move(filePath, newFilePath);
                WriteToEventLog($"Renamed file to: {newFilePath}", EventLogEntryType.Information, 6000);
            }
            catch (Exception ex)
            {
                WriteToEventLog($"Error renaming file: {ex.Message}", EventLogEntryType.Error, 7006);
                throw; // Volver a lanzar la excepción para capturarla en ProcessFiles
            }
        }
        /*Función que se ejecuta al parar el servicio*/
        protected override void OnStop()
        {
            try
            {
                // Detener todos los temporizadores individuales y registrar la carpeta que monitorizaban
                lock (timers)
                {
                    foreach (var timerInfo in timers.Values)
                    {
                        timerInfo.Timer.Stop();
                        WriteToEventLog($"Timer for folder {timerInfo.Folder} stopped.", EventLogEntryType.Information, 2002);
                    }
                }

                // Detener el temporizador de supervisión
                supervisionTimer?.Stop();

                WriteToEventLog("Service stopped.", EventLogEntryType.Information, 2001);
            }
            catch (Exception ex)
            {
                WriteToEventLog($"Error stopping service: {ex.Message}", EventLogEntryType.Error, 7007);
            }
        }
        /*Función para escribir los logs en el eventlog de windows*/
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
/*Clase para manejar toda la información de los timers*/
public class TimerInfo
{
    public Timer Timer { get; set; }
    public DateTime LastRunTime { get; set; }
    public string Folder { get; set; }
    public string Printer { get; set; }
}

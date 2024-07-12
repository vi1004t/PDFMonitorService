# PDF Monitor Service

PDF Monitor Service es una aplicaci�n de servicio de Windows que monitorea carpetas espec�ficas para archivos PDF y los imprime autom�ticamente utilizando una impresora configurada.

## Caracter�sticas

- Monitorea m�ltiples carpetas para archivos PDF.
- Imprime archivos PDF autom�ticamente utilizando impresoras configuradas.
- Maneja errores de impresi�n y evita el renombrado de archivos en caso de fallo.
- Registra eventos y errores en el Event Log de Windows.

## Requisitos

- .NET Framework 4.7.2 o superior.
- Windows 7 o superior.
- [PdfiumViewer](https://github.com/pvginkel/PdfiumViewer) (Paquete NuGet).

## Instalaci�n

1. Clona este repositorio en tu m�quina local:
   ```bash
   git clone https://github.com/tuusuario/pdf-monitor-service.git
   ```
2. Abre la soluci�n en Visual Studio.
- Ve a Tools > NuGet Package Manager > Package Manager Console.
- Ejecuta el siguiente comando:
```bash
dotnet restore
```
4. Compila la soluci�n

## Configuraci�n
1. Abre el archivo App.config en el proyecto y configura las secciones de folderPrinterMappings y appSettings seg�n tus necesidades.
2. Aqu� hay un ejemplo de configuraci�n:
```bash
<configuration>
  <configSections>
    <section name="folderPrinterMappings" type="System.Configuration.NameValueSectionHandler"/>
  </configSections>

  <folderPrinterMappings>
    <add key="C:\Path\To\Monitor1" value="Microsoft Print to PDF"/>
    <add key="C:\Path\To\Monitor2" value="Printer2"/>
    <!-- Agrega m�s duplas carpeta-impresora aqu� -->
  </folderPrinterMappings>
  <appSettings>
    <add key="EventSourceName" value="PDFMonitorServiceSource1" />
    <add key="IntervalInSeconds" value="3600" />
    <add key="FileSuffix" value="_processed" />
    <add key="RetryDelayInMilliseconds" value="1000" />
    <add key="PrintTimeoutInMilliseconds" value="10000" />
    <add key="EnablePrintRetry" value="true" />
    <add key="MaxPrintRetries" value="3" />
    <add key="MaxTimerRetries" value="0" />
    <add key="DebugMode" value="true"/>
  </appSettings>
</configuration>
```

## Uso
1. Ejecuta el servicio:
- Abre una terminal con privilegios de administrador y ejecuta los siguientes comandos:
```bash
sc create PDFMonitorService binPath= "C:\Path\To\PDFMonitorService.exe"
sc config PDFMonitorService start= auto
sc description PDFMonitorService "Servicio de impresi�n de partes de Siroper�a. Monitoriza las carpetas donde se generan los archivos PDF y los imprime."
sc start PDFMonitorService
```
2. Verifica los logs en el Event Viewer de Windows bajo PDFMonitorServiceLog.

## Licencia
Este proyecto est� bajo una licencia restrictiva. No puede ser utilizado, copiado o distribuido sin el permiso expl�cito del autor. Para obtener permiso, por favor contacta a vicentac+github@gmail.com
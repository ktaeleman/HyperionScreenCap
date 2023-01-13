using HyperionScreenCap.Config;
using log4net;
using log4net.Config;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Management;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace HyperionScreenCap
{
    internal static class Program
    {

        private static ILog LOG;

        static Program()
        {
        }

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("User32.dll")]
        public static extern bool SetProcessDpiAwarenessContext(int dpiFlag);

        [DllImport("SHCore.dll")]
        public static extern bool SetProcessDpiAwareness(PROCESS_DPI_AWARENESS awareness);

        public enum PROCESS_DPI_AWARENESS
        {
            Process_DPI_Unaware = 0,
            Process_System_DPI_Aware = 1,
            Process_Per_Monitor_DPI_Aware = 2
        }

        public enum DPI_AWARENESS_CONTEXT
        {
            DPI_AWARENESS_CONTEXT_UNAWARE = 16,
            DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = 17,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE = 18,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = 34
        }


           
        ///
        private static void SetDpiAwareness()
        {
            // Get windows version to set the correct dpi awareness context
            // This is needed for DXGI 1.5 OutputDuplicate1 to work
            var query = "SELECT * FROM Win32_OperatingSystem";
            var searcher = new ManagementObjectSearcher(query);
            var info = searcher.Get().Cast<ManagementObject>().FirstOrDefault();
            var version = info.Properties["Version"].Value.ToString();
            Version winVersion = new Version(version);
            // Windows 8.1 added support for per monitor DPI
            if ( winVersion >= new Version(6, 3, 0))
            {
                // Windows 10 creators update added support for per monitor v2
                if ( winVersion >= new Version(10, 0, 15063))
                {
                    SetProcessDpiAwarenessContext((int)DPI_AWARENESS_CONTEXT.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
                }
                else
                {
                    SetProcessDpiAwareness(PROCESS_DPI_AWARENESS.Process_Per_Monitor_DPI_Aware);
                }
            }
            else
            {
                SetProcessDPIAware();
            };
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            ConfigureLog4Net();
            LOG = LogManager.GetLogger(typeof(Program));
            LOG.Info("**********************************************************");
            LOG.Info("Application Startup. Logger Initialized.");
            LOG.Info("**********************************************************");

            // Set DPI awareness
            SetDpiAwareness();

            // Check if already running and exit if that's the case
            if (IsProgramRunning("hyperionscreencap", 0) > 1)
            {
                LOG.Error("Hyperion Screen Capture process already running.");
                try
                {
                    MessageBox.Show("HyperionScreenCap is already running!");
                    LOG.Info("Exiting application");
                    Environment.Exit(0);
                }
                catch (Exception)
                {
                    // ignored
                }
            }

            // Copy settings from previous version
            SettingsManager.CopySettingsFromPreviousVersion();

            // Migrate legacy settings
            SettingsManager.MigrateLegacySettings();

            // GitHub API requires TLS 1.2
            ConfigureSSL();

            MainForm _mainForm = new MainForm();
            Application.Run(_mainForm);
        }

        private static void ConfigureLog4Net()
        {
            log4net.GlobalContext.Properties["logFilePath"] = MiscUtils.GetLogDirectory() + Path.DirectorySeparatorChar + AppConstants.LOG_FILE_NAME;
            using ( Stream configStream = MiscUtils.GenerateStreamFromString(Resources.LogConfiguration) )
            {
                XmlConfigurator.Configure(configStream);
            }
        }

        private static int IsProgramRunning(string name, int runtime)
        {
            runtime += Process.GetProcesses().Count(clsProcess => clsProcess.ProcessName.ToLower().Equals(name));
            return runtime;
        }

        private static void ConfigureSSL()
        {
            // Adding TLS 1.1 & TLS 1.2
            // See: https://stackoverflow.com/questions/47269609/system-net-securityprotocoltype-tls12-definition-not-found
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol |= (SecurityProtocolType) (0x300 | 0xc00);
        }
    }
}
using System;
using System.Diagnostics;
using System.Security.Principal;
using Grapevine.Interfaces.Server;
using Grapevine.Server;
using Grapevine.Server.Attributes;
using Grapevine.Shared;
using log4net;

namespace HyperionScreenCap
{
    class ApiServer
    {
        private static readonly ILog LOG = LogManager.GetLogger(typeof(ApiServer));

        private MainForm _mainForm;
        private RestServer _server;

        public ApiServer(MainForm mainForm)
        {
            _mainForm = mainForm;
        }

        public void StartServer(string hostname, string port)
        {
            try
            {
                LOG.Info("Checking if ACL URL is reserved");
                if (!IsAclUrlReserved(hostname, port))
                {
                    LOG.Info("ACL URL not reserved. Attempting to reserve.");
                    ReserveAclUrl(hostname, port);
                }

                if ( _server == null )
                {
                    LOG.Info($"Starting API server: {hostname}:{port}");
                    _server = new RestServer
                    {
                        Host = hostname,
                        Port = port
                    };

                    var apiRoute = new Route(API);
                    _server.Router.Register(apiRoute);

                    _server.Start();
                    LOG.Info("API server started");
                }
            }
            catch ( Exception ex )
            {
                LOG.Error("Failed to start API server", ex);
            }
        }

        public void StopServer()
        {
            LOG.Info("Stopping API server");
            _server?.Stop();
            LOG.Info("API server stopped");
        }

        public void RestartServer(string hostname, string port)
        {
            LOG.Info("Restarting API server");
            StopServer();
            StartServer(hostname, port);
        }

        /// <summary>
        /// DO NOT RENAME THIS METHOD. The name is used in the reflection code above.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        [RestRoute(HttpMethod = HttpMethod.GET, PathInfo = "/API")]
        private IHttpContext API(IHttpContext context)
        {
            LOG.Info("API server command received");
            context.Response.ContentType = ContentType.TEXT;
            string responseText = "No valid API command received.";
            string command = context.Request.QueryString["command"] ?? "";
            string force = context.Request.QueryString["force"] ?? "false";

            if ( !string.IsNullOrEmpty(command) )
            {
                LOG.Info($"Processing API command: {command}");
                // Only process valid commands
                if ( command == "ON" || command == "OFF" )
                {

                    // Check for deactivate API between certain times
                    if ( SettingsManager.ApiExcludedTimesEnabled && force.ToLower() == "false" )
                    {
                        if ( (DateTime.Now.TimeOfDay >= SettingsManager.ApiExcludeTimeStart.TimeOfDay &&
                             DateTime.Now.TimeOfDay <= SettingsManager.ApiExcludeTimeEnd.TimeOfDay) ||
                            ((SettingsManager.ApiExcludeTimeStart.TimeOfDay > SettingsManager.ApiExcludeTimeEnd.TimeOfDay) &&
                             ((DateTime.Now.TimeOfDay <= SettingsManager.ApiExcludeTimeStart.TimeOfDay &&
                               DateTime.Now.TimeOfDay <= SettingsManager.ApiExcludeTimeEnd.TimeOfDay) ||
                              (DateTime.Now.TimeOfDay >= SettingsManager.ApiExcludeTimeStart.TimeOfDay &&
                               DateTime.Now.TimeOfDay >= SettingsManager.ApiExcludeTimeEnd.TimeOfDay))) )
                        {
                            responseText = "API exclude times enabled and within time range.";
                            LOG.Info($"Sending response: {responseText}");
                            context.Response.SendResponse(responseText);
                            return context;
                        }
                    }

                    _mainForm.ToggleCapture((MainForm.CaptureCommand) Enum.Parse(typeof(MainForm.CaptureCommand), command));
                    responseText = $"API command {command} completed successfully.";
                }

                if ( command == "STATE" )
                {
                    responseText = $"{_mainForm.CaptureEnabled}";
                }
            }
            else
            {
                LOG.Warn("API Command Empty / Invalid");
            }
            LOG.Info($"Sending response: {responseText}");
            context.Response.SendResponse(responseText);
            return context;
        }

        private string GetAclUrl(string hostname, string port)
        {
            return "http://" + hostname + ":" + port + "/";
        }

        private bool IsAclUrlReserved(string hostname, string port)
        {
            var aclUrl = GetAclUrl(hostname, port);
            ProcessStartInfo processStartInfo = new ProcessStartInfo
            {
                FileName = "netsh.exe",
                CreateNoWindow = true,
                UseShellExecute = false,
                Arguments = $"http show urlacl url={aclUrl}",
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true
            };
            LOG.Info($"Starting process: {processStartInfo.FileName} {processStartInfo.Arguments}");
            var process = Process.Start(processStartInfo);
            process.WaitForExit();
            var output = process.StandardOutput.ReadToEnd();
            /*
             * Sample output:
             *
             * ACL URL Not Reserved:
             * URL Reservations:
             * -----------------
             *
             * ACL URL Reserved:
             * URL Reservations:
             * -----------------
             *
             * Reserved URL            : http://+:9191/
             * User: DOMAIN\user
             * Listen: Yes
             * Delegate: No
             * SDDL: D:(A;;GX;;;S-1-5-21-566402754-1856570991-3730105997-1001)
             */
            return output.Contains(aclUrl);
        }

        private void ReserveAclUrl(string hostname, string port)
        {
            var aclUrl = GetAclUrl(hostname, port);
            var user = WindowsIdentity.GetCurrent().Name;
            ProcessStartInfo processStartInfo = new ProcessStartInfo
            {
                FileName = "netsh.exe",
                CreateNoWindow = true,
                UseShellExecute = true,
                Arguments = $"http add urlacl url={aclUrl} user={user}",
                WindowStyle = ProcessWindowStyle.Hidden,
                Verb = "runas",
            };
            LOG.Info($"Starting elevated process: {processStartInfo.FileName} {processStartInfo.Arguments}");
            var process = Process.Start(processStartInfo);
            process.WaitForExit();
        }
    }
}
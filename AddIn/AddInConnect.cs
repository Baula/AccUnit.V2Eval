using AccessCodeLib.AccUnit.V2Eval.Logging;
using AccessCodeLib.AccUnit.V2Eval.AddIn.Logging;
using Extensibility;
using Microsoft.Vbe.Interop;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.V2Eval.AddIn
{
    [ComVisible(true)]
    [GuidAttribute("57f7e5d7-cd64-4acb-bed8-71dd757844fc")]
    [ProgId("AccessCodeLib.AccUnit.V2Eval.AddIn.AddInConnect")]
    public class AddInConnect : IDTExtensibility2
    {
        private static readonly ILogger _staticLogger = new Logger();
        private IAddInLogger _logger;
        private Microsoft.Vbe.Interop.AddIn _addInInstance;
        private HostApplication _hostApplication;
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInstance, ref Array custom)
        {
            _logger = new AddInLogger();
            _logger.Log("OnConnection");
            _addInInstance = (Microsoft.Vbe.Interop.AddIn)addInInstance;
            try
            {
                _hostApplication = CreateHostApplication((VBE)application);
                _logger.Log("Application is " + _hostApplication.Name);
                _logger.LogVbProjects(_addInInstance.VBE);
            }
            catch (Exception xcp)
            {
                _logger.Log(xcp.ToString());
            }
        }

        private HostApplication CreateHostApplication(VBE vbe)
        {
            //var name = (string)application.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, application, new object[] { });

            return new HostApplication(vbe);
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _logger.Log("OnDisconnection");
        }

        public void OnStartupComplete(ref Array custom)
        {
            _logger.Log("OnStatupComplete");
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            _logger.Log("OnAddInsUpdate");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            _logger.Log("OnBeginShutdown");
        }

        [ComRegisterFunction]
        public static void RegisterClass(Type type)
        {
            _staticLogger.Log("RegisterClass: " + type.Name);
		}

        [ComUnregisterFunction]
        public static void UnregisterClass(Type type)
        {
            _staticLogger.Log("UnregisterClass: " + type.Name);
		}
    }
}

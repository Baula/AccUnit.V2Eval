using AccessCodeLib.AccUnit.V2Eval;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.V2Eval.Logging
{
    interface ILogger
    {
        void Log(string message);

        void LogVbProjects(VBE vbe);
    }

    class Logger : ILogger
    {
        public void Log(string message)
        {
            Debug.WriteLine("AccUnit.V2Eval: " + message);
        }

        public void LogVbProjects(VBE vbe)
        {
            foreach (VBProject vbProject in vbe.VBProjects)
            {
                LogVbProject(vbProject, isActiveVbProject: vbProject == vbe.ActiveVBProject);
            }
        }

        private void LogVbProject(VBProject vbProject, bool isActiveVbProject)
        {
            Log("Project " + vbProject.Name + ((isActiveVbProject) ? " (active)" : string.Empty));
            foreach (VBComponent vbComponent in vbProject.VBComponents)
            {
                LogVbComponent(vbComponent);
            }
        }

        private void LogVbComponent(VBComponent vbComponent)
        {
            Log(vbComponent.Name + " (" + vbComponent.Type.AsString() + ")");
        }
    }
}

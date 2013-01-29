using AccessCodeLib.AccUnit.V2Eval.AddIn;
using AccessCodeLib.AccUnit.V2Eval.Logging;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.V2Eval.AddIn.Logging
{
    interface IAddInLogger : ILogger
    {
        void LogVbProjects(VBE vbe);
    }

    class AddInLogger : Logger, IAddInLogger
    {
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

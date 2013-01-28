using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessCodeLib.AccUnit.V2Eval
{
    class HostApplication
    {
        private string _name;

        public HostApplication(VBE vbe)
        {
            var hostApplication = GetHostApplication(vbe);
            if (hostApplication != null)
                _name = GetPropertyValue<string>(hostApplication, "Name");
            else
                _name = "unknown host application";
        }

        private object GetHostApplication(VBE vbe)
        {
            foreach (VBProject vbProject in vbe.VBProjects)
            {
                var fileName = GetFileName(vbProject);
                if (fileName == null)
                    throw new Exception("You must save your file first.");
                var currentApplication = GetApplicationFromFileName(fileName);
                if (currentApplication != null)
                {
                    var appVbe = GetPropertyValue<VBE>(currentApplication, "VBE");
                    if (appVbe == vbe)
                        return currentApplication;
                }
            }
            return null;
        }

        private string GetFileName(VBProject vbProject)
        {
            try
            {
                return vbProject.FileName;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private object GetApplicationFromFileName(string fileName)
        {
            var directObject = GetObject(fileName);
            if (directObject == null)   // WinWord: VBProject = Normal.dot
                return null;
            string name;
            VBE vbe;
            if (GetNameAndVbeFromObject(directObject, out name, out vbe))
            {
                // The Access case
                return directObject;
            }
            var parentObject = GetPropertyValue<object>(directObject, "Parent");
            if (GetNameAndVbeFromObject(parentObject, out name, out vbe))
            {
                // The Excel case
                return parentObject;
            }
            var applicationObject = GetPropertyValue<object>(parentObject, "Application");
            if (GetNameAndVbeFromObject(applicationObject, out name, out vbe))
            {
                // The PowerPoint case
                return applicationObject;
            }
            throw new Exception("Unable to determine host application");
        }

        private bool GetNameAndVbeFromObject(object obj, out string name, out VBE vbe)
        {
            vbe = null;
            name = GetPropertyValueSafe<string>(obj, "Name");
            if (name == null)
                return false;
            vbe = GetPropertyValueSafe<VBE>(obj, "VBE");
            if (vbe == null)
                return false;
            return true;
        }

        private T GetPropertyValueSafe<T>(object obj, string propertyName)
        {
            try
            {
                return GetPropertyValue<T>(obj, propertyName);
            }
            catch (Exception)
            {
                return default(T);
            }
        }

        private object GetObject(string fileName)
        {
            try
            {
                return Marshal.BindToMoniker(fileName);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public string Name
        {
            get { return _name; }
        }

        private T GetPropertyValue<T>(object obj, string PropertyName)
        {
            return (T)obj.GetType().InvokeMember(PropertyName, BindingFlags.GetProperty, null, obj, null);
        }
    }
}

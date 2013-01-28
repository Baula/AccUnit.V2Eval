using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.V2Eval
{
    static class VbeExtensions
    {
        public static string AsString(this vbext_ComponentType type)
        {
            switch (type)
            {
                case vbext_ComponentType.vbext_ct_StdModule: return "Standard Module";
                case vbext_ComponentType.vbext_ct_ClassModule: return "Class Module";
                case vbext_ComponentType.vbext_ct_ActiveXDesigner: return "ActiveX Designer";
                case vbext_ComponentType.vbext_ct_Document: return "Document";
                case vbext_ComponentType.vbext_ct_MSForm: return "MSForm";
                default: return "(other type)";
            }
        }
    }
}

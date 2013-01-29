using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace AccessCodeLib.AccUnit.V2Eval.Logging
{
    public interface ILogger
    {
        void Log(string message);
    }

    public class Logger : ILogger
    {
        public void Log(string message)
        {
            Debug.WriteLine("AccUnit.V2Eval: " + message);
        }
    }
}

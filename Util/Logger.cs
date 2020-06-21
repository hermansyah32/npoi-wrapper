using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIWrapper.Util
{
    using Serilog;

    class Logger
    {
        private string methodName;

        public Logger(string methodName)
        {
            this.methodName = methodName;
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .CreateLogger();
        }

        public void Debug(string message)
        {
            Log.Debug(methodName + "=>" +  message);
        }

        public void Debug(Exception e, string message)
        {
            Log.Debug(e, methodName + "=>" + message);
        }

        public void Error(string message)
        {
            Log.Error(methodName + "=>" + message);
        }

        public void Error(Exception e, string message)
        {
            Log.Error(e, methodName + "=>" +  message);
        }
    }
}

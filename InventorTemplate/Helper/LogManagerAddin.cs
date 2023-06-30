using System;
using System.IO;
using System.Reflection;
using NLog.Config;
using NLog;

namespace InventorTemplate.Helper
{
    internal class LogManagerAddin
    {
        public static LogFactory Instance => _instance.Value;
        private static Lazy<LogFactory> _instance = new Lazy<LogFactory>(BuildLogFactory);

        private static LogFactory BuildLogFactory()
        {
            var basePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            if (basePath != null)
            {
                var configFilePath = Path.Combine(basePath, "nlog.config");

                var logFactory = new LogFactory();
                logFactory.Configuration = new XmlLoggingConfiguration(configFilePath, logFactory);
                logFactory.ThrowExceptions = true;
                return logFactory;
            }
            return null;
        }
    }
}
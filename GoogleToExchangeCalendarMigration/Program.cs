using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using log4net;
using log4net.Config;

namespace ConsoleApplication1
{
    class Program
    {
        protected static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static void Main(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();
            log.Info("STARTING CALENDAR MIGRATION......");
            
            GoogleCalendarApi gApi = new GoogleCalendarApi();
            gApi.WriteEventInformation();

            log.Debug("CALENDAR MIGRATION COMPLETE");
        }
    }
}

using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPO.ClientManager
{
    public static class LogManager
    {

        public static void WriteToFile(string text)
        {
            //string path = @"C:\\Update User Profile using AD\Logs";
            string path = ConfigurationManager.AppSettings["Path"];
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string logpath = path + "/Log_" + DateTime.Today.ToString("dd-MM-yyyy") + ".txt";
            using (StreamWriter writer = new StreamWriter(logpath, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }
    }
}

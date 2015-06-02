using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mailer.Class
{
    public class ConfigHelper
    {
        public string AktualnaSciezkaExcela
        {
            get
            {
                return ConfigurationManager.AppSettings["ExcelPlik"];
            }
        }

        public void PodmienUstawieniaSciezkiExcela(string value)
        {
            string appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configFile = System.IO.Path.Combine(appPath, "Mailer.exe.config");
            ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
            configFileMap.ExeConfigFilename = configFile;
            System.Configuration.Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
            config.AppSettings.Settings["ExcelPlik"].Value = value;
            config.Save();
        }
    }
}

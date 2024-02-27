using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.ComponentModel;
using System.IO;
using System.Xml.Linq;

namespace MAPacificReportUtility
{
    class UserSettings : INotifyPropertyChanged
    {
        const string Config_FileName = "MAPReportUtility.config";
        const string Path_Config_FileName = @"MAPReportUtility\MAPReportUtility.config";

        static UserSettings current;
        public static UserSettings Current
        {
            get
            {
                lock (typeof(UserSettings))
                {
                    if (current == null)
                    {
                        LoadSettings();
                    }
                    return current;
                }
            }
        }

        public static void Reset()
        {
            lock (typeof(UserSettings))
            {
                current = null;
            }
        }

        /// <summary>
        /// Correct the roaming settings file needed because sometimes the node "configSections" is missing from the config file.
        /// </summary>
        public static void CorrectRoamingSettingsFile()
        {
            const string NODE_NAME_CONFIGURATION = "configuration";
            const string NODE_NAME_CONFIGSECTIONS = "configSections";
            const string NODE_NAME_REPORTUTILITYSETTING = "ReportUtilitySettings";
            const string ADDIN_EXE_FILENAME = "MAPacificReportUtility.exe";

            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            //Use XDocument to load file and find the presence of the "configSections" node. If it does not exist, then add it to the file.
            XDocument xDoc = XDocument.Load(Path.Combine(appData, Path_Config_FileName));
            var configSectionsNode = xDoc.Element(NODE_NAME_CONFIGURATION).Element(NODE_NAME_CONFIGSECTIONS);

            if (configSectionsNode != null)
            {
                return;
            }
            var reportUtilitySettingsNode = xDoc.Element(NODE_NAME_CONFIGURATION).Element(NODE_NAME_REPORTUTILITYSETTING);
            var addInExeConfigFullFilename = AppDomain.CurrentDomain.BaseDirectory + ADDIN_EXE_FILENAME;
            var configDefault = ConfigurationManager.OpenExeConfiguration(addInExeConfigFullFilename);
            var xDocDefault = XDocument.Load(configDefault.FilePath);
            configSectionsNode = xDocDefault.Element(NODE_NAME_CONFIGURATION).Element(NODE_NAME_CONFIGSECTIONS);
            reportUtilitySettingsNode.AddBeforeSelf(configSectionsNode);
            xDoc.Save(Path.Combine(appData, Path_Config_FileName));
        }

        //Vista requires elevated priviliges in order to write to files stored in the C:\Program Files location.
        //To work around that, we will try to load the configuration file from the user's profile settings. If it does not exist, then
        //we will create a new one during run time when a save action occurs.
        public static void LoadSettings()
        {
            current = new UserSettings();

            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            ExeConfigurationFileMap exeMap = new ExeConfigurationFileMap();
            Configuration config = null;

            //load the config file from the user
            try
            {
                exeMap.ExeConfigFilename = Config_FileName;
                exeMap.RoamingUserConfigFilename = Path.Combine(appData, Path_Config_FileName);
                config = ConfigurationManager.OpenMappedExeConfiguration(exeMap, ConfigurationUserLevel.PerUserRoaming);
                //Fix for incompatible version of MapReportUtility.config file that is missing the <configsections> in the config file.
                //manually add that section into the file.
                //NOTE: For some unknown reason, the ConfigurationManager is not able to find the "configSections" node in the config file. This ends
                //up causing problems when using custom configuration sections. We will manually fix this in the CorrectRoamingSettingsFile method.
                ConfigurationSectionGroup configSections = config.GetSectionGroup("configSections");
                if (configSections == null)
                {
                    CorrectRoamingSettingsFile();
                    config = ConfigurationManager.OpenMappedExeConfiguration(exeMap, ConfigurationUserLevel.PerUserRoaming);
                }


                ReportUtilitySettingsSection reportSettingsSection = config.GetSection("ReportUtilitySettings") as ReportUtilitySettingsSection;
                if(reportSettingsSection == null)
                {
                    reportSettingsSection = new ReportUtilitySettingsSection();
                    reportSettingsSection.SectionInformation.AllowExeDefinition = ConfigurationAllowExeDefinition.MachineToRoamingUser;
                    config.Sections.Add("ReportUtilitySettings", reportSettingsSection);
                }

                if (config != null)
                {
                    current.ExcelDirectory = GetValue(config, "ExcelDirectory", string.Empty);
                    current.ExcelFileNamePrefix = GetValue(config, "ExcelFileNamePrefix", string.Empty);
                    current.DraftEmailPath = GetValue(config, "DraftEmailPath", string.Empty);
                }

            }
            catch (Exception)
            {
                // ToDo: log it somewhere
            }
        }

        public static void Save()
        {
            if (current != null)
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                ExeConfigurationFileMap exeMap = new ExeConfigurationFileMap();
                exeMap.ExeConfigFilename = Config_FileName;
                exeMap.RoamingUserConfigFilename = Path.Combine(appData, Path_Config_FileName);
                Configuration config = null;

                try
                {
                    config = ConfigurationManager.OpenMappedExeConfiguration(exeMap, ConfigurationUserLevel.PerUserRoaming);

                    if (config.GetSection("ReportUtilitySettings") == null)
                    {
                        ReportUtilitySettingsSection settings = new ReportUtilitySettingsSection();
                        settings.SectionInformation.AllowExeDefinition = ConfigurationAllowExeDefinition.MachineToRoamingUser;
                        config.Sections.Add("ReportUtilitySettings", settings);
                        config.Save(ConfigurationSaveMode.Minimal);
                    }


                    //clear contents before saving. The problem with using the appSettings in the config file is that if you use the
                    //same key name, then the value is inserted into the container and now you have two different values instead of just one.

                    ReportUtilitySettingsSection reportUtilitySection = config.GetSection("ReportUtilitySettings") as ReportUtilitySettingsSection;
                    reportUtilitySection.Settings.Clear();
                    reportUtilitySection.Settings.Add(new NameValueConfigurationElement("ExcelDirectory", current.ExcelDirectory));
                    reportUtilitySection.Settings.Add(new NameValueConfigurationElement("ExcelFileNamePrefix", current.ExcelFileNamePrefix));
                    reportUtilitySection.Settings.Add(new NameValueConfigurationElement("DraftEmailPath", current.DraftEmailPath));
                    config.Save(ConfigurationSaveMode.Modified);
                }
                catch (Exception e)
                {
                    // log it somewhere
                    System.Windows.Forms.MessageBox.Show(e.Message, "Application Settings Error", System.Windows.Forms.MessageBoxButtons.OK);
                }
            }
        }

        static T GetValue<T>(System.Configuration.Configuration config, string propertyName, T defaultValue)
        {
            try
            {
                ReportUtilitySettingsSection reportSettingsSection = config.GetSection("ReportUtilitySettings") as ReportUtilitySettingsSection;
                NameValueConfigurationElement val = reportSettingsSection.Settings[propertyName];
                if (val != null)
                {
                    if (val.Value != null)
                    {
                        return (T)Convert.ChangeType(val.Value, typeof(T));
                    }
                }
            }
            catch(Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK);
            }
            return defaultValue;
        }

        #region configurable settings

        private String excelDirectory;

        public String ExcelDirectory
        {
            get { return excelDirectory; }
            set
            {
                if (excelDirectory != value)
                {
                    excelDirectory = value;
                    OnPropertyChanged("ExcelDirectory");
                }
            }
        }

        //VI3596N Added new feature to store the excel file name prefix. This value is used
        //when generating individual excel files for each branch customer.
        private String excelFileNamePrefix = "";

        /// <summary>
        /// Property used for storing the file name prefix when generating individual customer excel reports.
        /// </summary>
        public String ExcelFileNamePrefix
        {
            get 
            {
                //default the return value to 
                if (excelFileNamePrefix != null && excelFileNamePrefix.Length == 0)
                {
                    return "Prepaid";
                }
                return excelFileNamePrefix; 
            }
            set
            {
                if (excelFileNamePrefix != value)
                {
                    excelFileNamePrefix = value;
                    OnPropertyChanged("ExcelFileNamePrefix");
                }
            }
        }

        private String draftEmailPath = "";
        public String DraftEmailPath
        {
            get { return draftEmailPath; }
            set
            {
                draftEmailPath = value;
                OnPropertyChanged("DraftEmailPath");
            }
        }

        #endregion

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }

        }
        #endregion
    }

    public class ReportUtilitySettingsSection : ConfigurationSection
    {
        /// <summary>
        /// name of the section in the config file
        /// </summary>
        /// 
        public const string SectionName = "ReportUtilitySettingsSection";

        [ConfigurationProperty("", IsDefaultCollection = true)]
        public NameValueConfigurationCollection Settings
        {
            get
            {
                return (NameValueConfigurationCollection)base[""];
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;
using Infragistics.Win.UltraWinGrid;
using Infragistics.Win;

namespace MAPacificReportUtility.forms
{
    public partial class SubClientIDPreferencesForm : Form
    {
        public SubClientIDPreferencesForm()
        {          
            InitializeComponent();
            ultraGridBranchInfo.DisplayLayout.AddNewBox.Hidden = false;
            //change the prompt of on the add new box
            ultraGridBranchInfo.DisplayLayout.AddNewBox.Prompt = "Add a new row";
            ultraGridBranchInfo.DisplayLayout.Bands[0].AddButtonCaption = "New Branch";
            _reportclientBranchList = new BindingList<ClientBranch>();

            ultraGridWireConf.DisplayLayout.AddNewBox.Hidden = false;
            ultraGridWireConf.DisplayLayout.AddNewBox.Prompt = "Add a new PCRNO";
            ultraGridWireConf.DisplayLayout.Bands[0].AddButtonCaption = "New PRCNO";

            myDContainer = new DataContainer();
            LoadConfiguration();

        }

        public SubClientIDPreferencesForm(BindingList<ClientBranch> inBindingList)
        {
            InitializeComponent();          
            ultraGridBranchInfo.DisplayLayout.AddNewBox.Hidden = false;
            //change the prompt of on the add new box
            ultraGridBranchInfo.DisplayLayout.AddNewBox.Prompt = "Add a new row";
            ultraGridBranchInfo.DisplayLayout.Bands[0].AddButtonCaption = "New Branch";
            _reportclientBranchList = inBindingList;

            ultraGridWireConf.DisplayLayout.AddNewBox.Hidden = false;
            ultraGridWireConf.DisplayLayout.AddNewBox.Prompt = "Add a new PCRNO";
            ultraGridWireConf.DisplayLayout.Bands[0].AddButtonCaption = "New PRCNO";
            myDContainer = new DataContainer();

            LoadConfiguration();
        }

        public struct ExcelSummaryType
        {
            public static string ALL
            {
                get { return "ALL"; }
            }

            public static string GPR_Gift_Summary_Prepaid_Total
            {
                get { return "GPR and Gift Summary and Prepaid Card Total"; }
            }

            public static string GPR_Gift_Details
            {
                get { return "GPR and Gift Details"; }
            }

        }
        #region  Dictionary and functions for ClientBranches..

        public bool syncUpdated = false;

        private BindingList<ClientBranch> _reportclientBranchList;
        public BindingList<ClientBranch> ReportClientBranchList
        {
            get { return _reportclientBranchList; }
            set
            {
                if(value != null)
                    _reportclientBranchList = value;
            }
        }

        //used to store Client Branch info retrieved from file
        private BindingList<ClientBranch> _fileclientBranchList;

        private DataContainer myDContainer;
        public DataContainer MyDContainer
        {
            get
            {
                return myDContainer;
            }
        }

      

        private string _configDirectory;
        /// <summary>
        /// Configuration Directory
        /// </summary>
        public string ConfigDirectory
        {
            get { return _configDirectory; }
            set
            {
                if (value != null)
                {
                    _configDirectory = value;
                }
            }
        }


        public void LoadConfiguration()
        {
            //Lets look for the file in the application launch directory first. If it does not exist then
            //we will look in the specialfolder location.
            string file = Path.Combine(Application.StartupPath, @"BranchInfo.xml");

            if (File.Exists(file))
            {
                //VS4229 support serializing both ClientBranch and string bindinglist.
                XmlSerializer deserializer = new XmlSerializer(typeof(DataContainer));
                TextReader textReader = new StreamReader(file);
                myDContainer = (DataContainer)deserializer.Deserialize(textReader);
                _fileclientBranchList = myDContainer.ClientBranchList;
                pRCNOItemBindingSource.DataSource = myDContainer.WireConfTableList;

                textReader.Close();
                SyncBindinglist();
                if (_reportclientBranchList.Count == 0)
                    ultraGridBranchInfo.DataSource = _fileclientBranchList;
                else
                    ultraGridBranchInfo.DataSource = _reportclientBranchList;

                //exit out.
                return;
            }


            
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            file = Path.Combine(appData, @"MAPReportUtility\BranchInfo.xml");

            if(File.Exists(file))
            {
                //VS4229 support serializing both ClientBranch and string bindinglist.
                XmlSerializer deserializer = new XmlSerializer(typeof(DataContainer));
                TextReader textReader = new StreamReader(file);
                if (textReader.Peek() > 0)
                {
                    myDContainer = (DataContainer)deserializer.Deserialize(textReader);
                    _fileclientBranchList = myDContainer.ClientBranchList;
                    pRCNOItemBindingSource.DataSource = myDContainer.WireConfTableList;
                    SyncBindinglist();
                }
                textReader.Close();
                //SyncBindinglist();
                if (_reportclientBranchList.Count == 0)
                    ultraGridBranchInfo.DataSource = _fileclientBranchList;
                else
                    ultraGridBranchInfo.DataSource = _reportclientBranchList;

            }

            if (!File.Exists(file))
            {
                //VS4229 if the file does not exist, then we need to populate the bindinglist with some default
                //values.
                ultraGridBranchInfo.DataSource = _reportclientBranchList;
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC208"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC339"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC347"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC451"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC564"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC586"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC600"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC603"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC607"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC615"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC665"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC692"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC765"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC766"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC823"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC828"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC966"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC968"));
                //VS6159 Add new numbers to the default list..
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC211"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC558"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC580"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC818"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC819"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC885"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC887"));
                myDContainer.WireConfTableList.Add(new PRCNOItem("PRC975"));
                pRCNOItemBindingSource.DataSource = myDContainer.WireConfTableList;
                return;
            }
        }

        public void SaveConfiguration()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string file = Path.Combine(appData, @"MAPReportUtility\BranchInfo.xml");

            if (!File.Exists(file))
            {
                try
                {
                    //Create the File and close the stream down so that it can be used later.
                    FileStream fs = File.Create(file);
                    fs.Close();
                }
                catch (DirectoryNotFoundException dex)
                {
                    Directory.CreateDirectory(Path.Combine(appData, @"MAPReportUtility"));
                }
            }

            //no data to save so skip the save and return.
            if (_reportclientBranchList.Count == 0 && _fileclientBranchList.Count == 0)
                return;

            XmlSerializer serializer = new XmlSerializer(typeof(DataContainer));
            TextWriter textWriter = new StreamWriter(file);
            if (_reportclientBranchList.Count > 0)
            {
                myDContainer.ClientBranchList = _reportclientBranchList;
            }
            else
            {
                myDContainer.ClientBranchList = _fileclientBranchList;
            }

            serializer.Serialize(textWriter, myDContainer);
            textWriter.Close();
        }

        #endregion

        /// <summary>
        /// Synchronize the BindingList that was retrieved from the BranchInfo.xml file into the one obtained from parsing the report.
        /// </summary>
        public void SyncBindinglist()
        {
            if (_fileclientBranchList.Count > 0 && _reportclientBranchList.Count > 0)
            {
                int index = -1;
                foreach (ClientBranch mBranch in _reportclientBranchList)
                {
                    index = _fileclientBranchList.IndexOf(mBranch);
                    if (index > -1)
                    {
                        mBranch.Category = _fileclientBranchList[index].Category;
                        mBranch.RelationalParentSummary = _fileclientBranchList[index].RelationalParentSummary;
                        //if default value for the excel summary is blank then default it to ALL
                        if (_fileclientBranchList[index].ExcelSummary.Length == 0)
                        {
                            mBranch.ExcelSummary = ClientBranch.ExcelSummaryType.ALL;
                        }
                        else
                        {
                            mBranch.ExcelSummary = _fileclientBranchList[index].ExcelSummary;
                        }
                        //VS3531 Update the Name from the BranchInfo.xml file into the current branchlist.
                        mBranch.Name = _fileclientBranchList[index].Name;
                        //VS4596 update the email address
                        mBranch.EmailAddress = _fileclientBranchList[index].EmailAddress;
                    }
                }

                if (_fileclientBranchList.Count != _reportclientBranchList.Count)
                {
                    //user added new rows of data. Need to include those into the reportbranch list.
           //         for (int i = _reportclientBranchList.Count; i < _fileclientBranchList.Count; i++)
                    //VS4583 The syncing up of report client branch and the client branch from the branchInfo.xml file did not work
                    //correctly in the previous logic setting. This time around go through the list and check for its presence in the
                    //report list, if not then add it into the report list.
                    for(int i = 0; i < _fileclientBranchList.Count; i++)
                    {
                        if(_reportclientBranchList.Contains(_fileclientBranchList[i]))
                            continue;
                        _reportclientBranchList.Add(_fileclientBranchList[i]);
                    }

                }

                ultraGridBranchInfo.DataSource = _reportclientBranchList;
                syncUpdated = true;
            }
            else  //edge case to handle populating the fileclientbranchlist with the default ExcelSummary of ALL.
            {     //Implementing this new feature in phase2 requires assigning a default value to this property that exist in the BranchInfo.xml file.
                if (_fileclientBranchList.Count > 0)
                {
                    for (int i = 0; i < _fileclientBranchList.Count; i++)
                    {
                        if (_fileclientBranchList[i].ExcelSummary.Length == 0)
                        {
                            _fileclientBranchList[i].ExcelSummary = ExcelSummaryType.ALL;
                        }
                    }
                }
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            SaveConfiguration();
        }

        private void ultraGridBranchInfo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            //VS4229 also add support to delete data from branchinfo.xml
            UltraGridLayout layout = e.Layout;
            UltraGridOverride ov = layout.Override;
            ov.AllowDelete = DefaultableBoolean.True;

            foreach ( Infragistics.Win.UltraWinGrid.UltraGridBand band in e.Layout.Bands)
            {
                band.Columns[1].PerformAutoResize(Infragistics.Win.UltraWinGrid.PerformAutoSizeType.AllRowsInBand);
                band.Columns[2].PerformAutoResize(Infragistics.Win.UltraWinGrid.PerformAutoSizeType.AllRowsInBand);
            }

        }

        private void ultraGridWireConf_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            //vs4229 also add support to delete records from the ultragrid.
            UltraGridLayout layout = e.Layout; 
            UltraGridOverride ov = layout.Override;
            ov.AllowDelete = DefaultableBoolean.True;
        }
    }
}

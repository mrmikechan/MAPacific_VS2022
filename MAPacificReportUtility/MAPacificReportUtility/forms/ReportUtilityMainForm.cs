using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinDataSource;
using Infragistics.Win.UltraWinGrid;
using MAPacificReportUtility.excel;
using System.Text.RegularExpressions;
using System.IO;
using OvDotNet;

/*
 * Change   Date    Name    Details
 * -------------------------------------------------------------------------------------------------------------------------------------------
 * $d01     033012  MC      Added delegates to get around cross threading gui updates.
 * $d02     033012  MC      Added Debug Traceline to aid in debugging tool.
 */


namespace MAPacificReportUtility.forms
{
    public partial class ReportUtilityMainForm : Form
    {
        //delegates used for cross thread communication to gui components...
        delegate void LogEventCallback(string text);
        delegate void TextboxReportDateCallback(string date);
        delegate void EnableGUICallback();
        delegate void ShowEditFormCallback();

        #region member properties/data

        const int PrePaidCardsTotal = 0;
        const int GPRSummary = 1;
        const int GiftCardSummary = 2;
        const int GPRBranchDetail = 3;
        const int GiftBranchDetail = 4;
        const int IgnoreReport = 5;
        const int UmbrellaReport = 6;
        const int UnrecognizedID = 7;

        //VS4142 Modification to work with Changes to VisaDPS Report
        //Contents of reports are changing and being split into seperate GPR and GIFT reports with their own BIN
        //number.
        //MAP GPR BIN#:  99586001
        //MAP GIFT BIN#: 99586002
        //VS4583 Added private Banks to the fold using Beken.
        //BEKEN GPR BIN#: 99818001
        //Beken GIFT GIFT BIN#:     unknown at this time and defaulted to 0000000.
        const string MAP_GPR_BIN = "99586001";
        const string MAP_GIFT_BIN = "99586002";
        const string BEKEN_GPR_BIN = "99818001";
        const string BEKEN_GIFT_BIN = "00000000";

        BindingList<ClientBranch> _prepaidcardtotalsummary;  //reports with subclientID code suffix 000
        BindingList<ClientBranch> _gprsummary;               //reports with subClientID code suffix 100   
        BindingList<ClientBranch> _giftsummary;              //reports with subClientID code suffix 100
        BindingList<ClientBranch> _gprdetail;                //reports with subClientID code suffix 1NN
        BindingList<ClientBranch> _giftdetail;               //reports with subClientID code suffix 1NN
        BindingList<ClientBranch> _ignorereport;
        BindingList<ClientBranch> _unrecognizedid;

        //used for storage into BranchInfo.xml file and also for verification of unrecognized branch info.
        BindingList<ClientBranch> ClientBranchList;
        //Contains all the clientbranches after parsing report...
        List<ClientBranch> ExportExcelList;
        //VS4731 and VS4732
        List<ClientBranch> ExportExcelCardActivityList;

        //Calculated Pre Paid Cards Total Summary
        List<TransactionType> CalculatedPrePaidTotalList;

        MAPacificExportExcel mExcel;
        GlobalSummaryBranch mGlobalSummaryData;

        private String visaReportRunDate = "";

        ProcessVisaDPSReport _processVisaReport;
        ProcessFinancialSummaryReport _processFinSumReport;
        BackgroundWorker bgwVisaDPS;
        BackgroundWorker bgwFinancialSummary;
        //VS4200 Process DPS Wire Confirmation Report
        ProcessWireConfirmationReport _processWireReport;
        BackgroundWorker bgwDPWireConfReport;

        static readonly object countLock = new object();

        UltraGrid mFinancialSummaryReportGrid;

        //flag to turn on/off Trace statements
        bool debug = false;
        Stopwatch sw;

        //VS4229 Implement capability to save Credit Union to BranchInfo.xml so that we know which
        //CU to write out to the Wire Confirmation Total column.
        DataContainer myDataContainer = null;
        #endregion

        public ReportUtilityMainForm()
        {
            InitializeComponent();
            //in IDE this value keeps getting reset to false while debugging...
            //ovApi.AutoStartOrAttach = true;
            //ovApi.OutputTraceMessages = true;
            //ovApi = new OvDotNet.OvDotNetApi(this.components);
            //ovApi.AutoUpdateFields = true;
            //ovApi.DataCaptureAccumulate = false;
            //ovApi.AutoStartOrAttach = true;
            //ovApi.EventFilterPeriod = 75;
            //ovApi.InvokeThough = this;
            //ovApi.NewOvInstance = false;
            //ovApi.SessionFilename = null;
            //ovApi.SimKeyboardState = true;
            //ovApi.Tag = null;
            //ovApi.SessionStateChanged += new OvDotNet.OvSessionStateEventHandler(this.ovApi_SessionStateChanged);
            //ovApi.KeyboardStateChanged += new OvDotNet.OvKeyboardStateEventHandler(this.ovApi_KeyboardStateChanged);
            //ovApi.DataReceived += new OvDotNet.OvDataReceivedEventHandler(this.ovApi_DataReceived);
            //ovApi.ScreenModeChanged += new OvDotNet.OvScreenModeEventHandler(this.ovApi_ScreenModeChanged);
      //      ovApi.DataFieldsUpdate += new OvDotNet.OvDataUpdateEventHandler(this.ovApi_DataFieldsUpdate);
      //      ovApi.EmulationTypeActive += new OvDotNet.OvEmulationTypeEventHandler(this.ovApi_EmulationTypeActive);
      //      ovApi.HotLinkToOv();

            sw = new Stopwatch();
            _processVisaReport = new ProcessVisaDPSReport(ovApi);
            _processVisaReport.FunctionKey = "PF8"; //PF8 key for paging to next page

            _processVisaReport.BranchChanged += new ProcessVisaDPSReport.BranchChangedEventHandler(_processVisaReport_BranchChanged);
            _processVisaReport.EndofReport += new ProcessVisaDPSReport.EndofReportEventHandler(_processVisaReport_EndofReport);
            _processVisaReport.ReportDate += new ProcessVisaDPSReport.ReporDateEventHandler(_processVisaReport_ReportDate);
            //initialize the different list buckets to hold the branch info...
            _prepaidcardtotalsummary = new BindingList<ClientBranch>();
            _gprsummary = new BindingList<ClientBranch>();
            _giftsummary = new BindingList<ClientBranch>();
            _gprdetail = new BindingList<ClientBranch>();
            _giftdetail = new BindingList<ClientBranch>();
            _ignorereport = new BindingList<ClientBranch>();
            _unrecognizedid = new BindingList<ClientBranch>();
            ClientBranchList = new BindingList<ClientBranch>();
            ExportExcelList = new List<ClientBranch>();
            CalculatedPrePaidTotalList = new List<TransactionType>();
            mGlobalSummaryData = new GlobalSummaryBranch();

            _processFinSumReport = new ProcessFinancialSummaryReport(ovApi);
            _processFinSumReport.FunctionKey = "PF8"; //PF8 key for paging to next page
            _processFinSumReport.EndofReport += new ProcessFinancialSummaryReport.EndofReportEventHandler(_processFinSumReport_EndofReport);

            //VS3623 Report utility is pausing intermitently. Need to rework the threading logic..
            //background worker thread to offload the processing of the visa report.
            bgwVisaDPS = new BackgroundWorker();
            bgwVisaDPS.WorkerSupportsCancellation = true;
            bgwVisaDPS.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);
            bgwVisaDPS.DoWork += new DoWorkEventHandler(bgw_DoWork);

            bgwFinancialSummary = new BackgroundWorker();
            bgwFinancialSummary.WorkerSupportsCancellation = true;
            bgwFinancialSummary.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgwFinancialSummary_RunWorkerCompleted);
            bgwFinancialSummary.DoWork += new DoWorkEventHandler(bgwFinancialSummary_DoWork);

            _processWireReport = new ProcessWireConfirmationReport(ovApi);
            _processWireReport.EndofReport += new ProcessWireConfirmationReport.EndofReportEventHandler(_processWireReport_EndofReport);
            bgwDPWireConfReport = new BackgroundWorker();
            bgwDPWireConfReport.WorkerSupportsCancellation = true;
            bgwDPWireConfReport.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgwDPWireConfReport_RunWorkerCompleted);
            bgwDPWireConfReport.DoWork += new DoWorkEventHandler(bgwDPWireConfReport_DoWork);

            FontData fd = ultraGridExcel.DisplayLayout.Appearance.FontData;
            fd.Name = "Arial";
            ultraGridExcel.InitializeLayout += new InitializeLayoutEventHandler(ultraGridExcel_InitializeLayout);

            FontData fd2 = ultraGridGlobalSummary.DisplayLayout.Appearance.FontData;
            fd2.Name = "Arial";
            ultraGridGlobalSummary.InitializeLayout += new InitializeLayoutEventHandler(ultraGridGlobalSummary_InitializeLayout);
        }

        /// <summary>
        /// InitializeLayout event allows us to preset the layout characteristics of the text in the cells for these columns. In instance
        /// we are horizontally aligning the text within a cell.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridGlobalSummary_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Bands[0].Columns["ClientID"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[1].Columns["TransactionCount"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[1].Columns["TransactionAmount"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[1].Columns["DBCR1"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[1].Columns["FeeAmount"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[1].Columns["DBCR2"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[1].Columns["TotalAmount"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[1].Columns["DBCR3"].CellAppearance.TextHAlign = HAlign.Center;

            e.Layout.Bands[2].Columns["FundsPoolBalance"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[2].Columns["DBCR"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[2].Columns["AccountsReported"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[2].Columns["OpenAccounts"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[2].Columns["ClosedAccounts"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[2].Columns["VoidedAccounts"].CellAppearance.TextHAlign = HAlign.Center;

        }

        /// <summary>
        /// MAPacific wants to horizontal align the text under ClientID for PrepaidIport worksheet.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcel_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Bands[0].Columns["ClientID"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["Name"].PerformAutoResize(Infragistics.Win.UltraWinGrid.PerformAutoSizeType.AllRowsInBand);
        }


       

        #region ovAPI events


        private void ovApi_DataFieldsUpdate(object sender, EventArgs e)
        {

        }

        private void ovApi_DataReceived(object sender, EventArgs e)
        {

        }

        private void ovApi_EmulationTypeActive(object sender, OvDotNet.OvEmulationType e)
        {

        }

        private void ovApi_ScreenModeChanged(object sender, OvDotNet.OvScreenModeArgs e)
        {

        }

        private void ovApi_SessionStateChanged(object sender, OvDotNet.OvSessionStateArgs e)
        {
            LogEvent(e.ConnectionState.ToString() + " " + e.Message);
            ActionStatus.Text = "OV Link" + e.ConnectionState.ToString();
            if (e.ConnectionState == OvDotNet.OvSessionState.Disconnected)
            {
                disableGUI();
                //VS3628 Cancel background worker if we are disconnected!
                lock (countLock)
                {
                    if (bgwVisaDPS.IsBusy)
                    {
                        TraceLine("Session Disconnected - cancel: " );
                        bgwVisaDPS.CancelAsync();
                        System.Threading.Monitor.Pulse(countLock);
                    }
                }
                LogEvent("OV Disconnected");
            }

            if (e.ConnectionState == OvDotNet.OvSessionState.Online)
            {
                enableGUI();
                LogEvent("OV Online");
            }
        }

        //VS3623 Added Monitor and object Lock for threading purpose to prevent dead lock or in our case pausing of the tool.
        //VS360 placing the parsing logic of the report into background thread.
        void ovApi_KeyboardStateChanged(object sender, OvDotNet.OvKeyboardStateArgs e)
        {
            //string state = "";
            //switch (e.State)
            //{
            //    case OvDotNet.OvKeyboardState.Locked:
            //        state = "Locked";
            //        break;
            //    case OvDotNet.OvKeyboardState.UnLocked:
            //        state = "Unlocked";
            //        break;
            //    case OvDotNet.OvKeyboardState.UserForcedUnlock:
            //        state = "UserForceUnlock";
            //        break;
            //}
            //System.Diagnostics.Trace.WriteLine("KeyboarSateChanged ThreadID: " + System.Threading.Thread.CurrentThread.ManagedThreadId + " State: " + state);

            if (e.State == OvDotNet.OvKeyboardState.UnLocked)
            {
                //VS3623
                lock (countLock)
                {
                    TraceLine("Unlock - Monitor Pulse called from threadID: " );
                    
                    System.Threading.Monitor.Pulse(countLock);
                }
            }
        }
        #endregion

        /// <summary>
        /// Log messages so that it can get displayed in the tool strip panel.
        /// </summary>
        /// <param name="text"></param>
        private void LogEvent(string text)
        {
            if (this.textBoxReportDate.InvokeRequired) //Check to see if we have cross threading access issue...
            {
                LogEventCallback lcb = new LogEventCallback(LogEvent);
                this.Invoke(lcb, new object[] { text });
            }
            else
            {
                toolStripStatusLabel2.Text = text;
            }
        }

        /// <summary>
        /// Disable all the GUI buttons so that users can't click on them.
        /// </summary>
        private void disableGUI()
        {
            buttonProcess.Enabled = false;
            buttonEdit.Enabled = false;
            buttonGenerate.Enabled = false;
            buttonGenerateClientReports.Enabled = false;
            buttonProcFinSummary.Enabled = false;
            buttonProcessWireConfReport.Enabled = false;
            btn_emailDraft.Enabled = false;
        }

        /// <summary>
        /// Enable all the GUI buttons so that users can click on them.
        /// </summary>
        private void enableGUI()
        {
            if (this.buttonProcess.InvokeRequired)
            {
                EnableGUICallback ecb = new EnableGUICallback(enableGUI);
                this.Invoke(ecb);
            }
            else
            {
                buttonProcess.Enabled = true;
                buttonEdit.Enabled = true;
                buttonGenerate.Enabled = true;
                buttonGenerateClientReports.Enabled = true;
                buttonProcFinSummary.Enabled = true;
                buttonProcessWireConfReport.Enabled = true;
                btn_emailDraft.Enabled = true;
            }
        }

        /// <summary>
        /// Check the values loaded from BranchInfo.xml and update local tables with those data...
        /// </summary>
        /// <returns></returns>
        private void MergeClientBranchInfo()
        {
            SubClientIDPreferencesForm mForm = new SubClientIDPreferencesForm(ClientBranchList);
            if (mForm.syncUpdated)
            {
                UpdateDataTables(mForm.ReportClientBranchList);
            }

            //vs4229 RETRIEVE THE information from the file and update that into the processWireReport object.
            if (mForm.MyDContainer != null)
            {
                myDataContainer = mForm.MyDContainer;
                if (myDataContainer.WireConfTableList.Count > 0)
                {
                    _processWireReport.SetDataContainer(myDataContainer);
                }
            }
        }

        /// <summary>
        /// Calculated the data within the Pre Paid Card Total table and sum those values
        /// up so that we can compare them with the Global summary data that is at the tail end
        /// of the VisaDPS report.  This provides us a quick check to see if the data matches.
        /// </summary>
        public void AddPrepaidCardTotalSummary()
        {
            //Loads / FI Funds Transfer 
            //Reloads / FI Funds Transfer 
            //Unloads / FI Funds Transfer  
            //Reloads / By  - Pass 
            //Reloads / Merchant ACQ (not in attached example but can be) 
            //Unloads / Merchant ACQ (not in attached example but can be) 
            //Unloads / By - Pass

            TransactionType loadsFIFunds = new TransactionType();
            loadsFIFunds.Transaction = TransactionType.TransactionOption.LOADS_FI_FUNDS_TRANSFER;
            TransactionType reloadsFIFunds = new TransactionType();
            reloadsFIFunds.Transaction = TransactionType.TransactionOption.RELOADS_FI_FUNDS_TRANSFER;
            TransactionType unloadsFIFunds = new TransactionType();
            unloadsFIFunds.Transaction = TransactionType.TransactionOption.UNLOADS_FI_FUNDS_TRANSFER;
            TransactionType unloadsByPass = new TransactionType();
            //VS3594 Add unloads/by-pass
            unloadsByPass.Transaction = TransactionType.TransactionOption.UNLOADS_BYPASS;
            TransactionType reloadsByPass = new TransactionType();
            reloadsByPass.Transaction = TransactionType.TransactionOption.RELOADS_BYPASS;
            TransactionType reloadsMerchACQ = new TransactionType();
            reloadsMerchACQ.Transaction = TransactionType.TransactionOption.RELOADS_MERCHANT_ACQ;
            TransactionType unloadsMerchACQ = new TransactionType();
            unloadsMerchACQ.Transaction = TransactionType.TransactionOption.UNLOADS_MERCHANT_ACQ;
            TransactionType totalLoadsUnloads = new TransactionType();
            totalLoadsUnloads.Transaction = TransactionType.TransactionOption.TOTAL_LOAD_UNLOAD_ACTIVITY;
            TransactionType manualAdj = new TransactionType();
            manualAdj.Transaction = TransactionType.TransactionOption.MANUAL_ADJUSTMENT;
            TransactionType loadsMerchPOSFunding = new TransactionType();
            loadsMerchPOSFunding.Transaction = TransactionType.TransactionOption.LOADS_MERCH_POS_FUNDING;
            //VS4637 Add new TransactionTypes per Olrando
            TransactionType loadsByPass = new TransactionType();
            loadsByPass.Transaction = TransactionType.TransactionOption.LOADS_BYPASS;
            TransactionType loadsMerchACQ = new TransactionType();
            loadsMerchACQ.Transaction = TransactionType.TransactionOption.LOADS_MERCHANT_ACQ;

            foreach (ClientBranch mBranch in _prepaidcardtotalsummary)
            {
                if (mBranch.TransactionList.Count > 0)
                {
                    decimal value;

                    //do not add data from summary branches that a relational parent summary value.
                    //The parent summary already contains the data for it and therefore we would end up counting the
                    //value twice.
                    if (mBranch.RelationalParentSummary.Length > 0)
                        continue;

                    foreach (TransactionType mType in mBranch.TransactionList)
                    {
                        try
                        {
                            if (mType.Transaction.Equals(TransactionType.TransactionOption.LOADS_FI_FUNDS_TRANSFER))
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = loadsFIFunds.TransactionCount + mType.TransactionCount;
                                loadsFIFunds.TransactionCount = value;
                                //calculate Transaction Amount
                                if (loadsFIFunds.DBCR1.Equals(mType.DBCR1))
                                    value = loadsFIFunds.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = loadsFIFunds.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    loadsFIFunds.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                loadsFIFunds.TransactionAmount = value;
                                //calculate Fee Amount
                                value = loadsFIFunds.FeeAmount + mType.FeeAmount;
                                loadsFIFunds.FeeAmount = value;
                                //calculate Total Amount
                                if (loadsFIFunds.DBCR2.Equals(mType.DBCR2))
                                    value = loadsFIFunds.TotalAmount + mType.TotalAmount;
                                else
                                    value = loadsFIFunds.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    loadsFIFunds.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                loadsFIFunds.TotalAmount = value;
                            }
                            else if (mType.Transaction.Equals(TransactionType.TransactionOption.RELOADS_FI_FUNDS_TRANSFER))
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = reloadsFIFunds.TransactionCount + mType.TransactionCount;
                                reloadsFIFunds.TransactionCount = value;
                                //calculate Transaction Amount
                                if (reloadsFIFunds.DBCR1.Equals(mType.DBCR1))
                                    value = reloadsFIFunds.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = reloadsFIFunds.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    reloadsFIFunds.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                reloadsFIFunds.TransactionAmount = value;
                                //calculate Fee Amount
                                value = reloadsFIFunds.FeeAmount + mType.FeeAmount;
                                reloadsFIFunds.FeeAmount = value;
                                //calculate Total Amount
                                if (reloadsFIFunds.DBCR2.Equals(mType.DBCR2))
                                    value = reloadsFIFunds.TotalAmount + mType.TotalAmount;
                                else
                                    value = reloadsFIFunds.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    reloadsFIFunds.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                reloadsFIFunds.TotalAmount = value;
                            }
                            else if (mType.Transaction.Equals(TransactionType.TransactionOption.LOADS_MERCH_POS_FUNDING))
                            {
                                //calculate the transaction count
                                value = 0;
                                value = loadsMerchPOSFunding.TransactionCount + mType.TransactionCount;
                                loadsMerchPOSFunding.TransactionCount = value;
                                //calculate Transaction Amount
                                if (loadsMerchPOSFunding.DBCR1.Equals(mType.DBCR1))
                                    value = loadsMerchPOSFunding.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = loadsMerchPOSFunding.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    loadsMerchPOSFunding.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                loadsMerchPOSFunding.TransactionAmount = value;
                                //calculate Fee Amount
                                value = loadsMerchPOSFunding.FeeAmount + mType.FeeAmount;
                                loadsMerchPOSFunding.FeeAmount = value;
                                //calculate Total Amount
                                if (loadsMerchPOSFunding.DBCR2.Equals(mType.DBCR2))
                                    value = loadsMerchPOSFunding.TotalAmount + mType.TotalAmount;
                                else
                                    value = loadsMerchPOSFunding.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    loadsMerchPOSFunding.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                loadsMerchPOSFunding.TotalAmount = value;
                            }

                            else if (mType.Transaction.Equals(TransactionType.TransactionOption.UNLOADS_FI_FUNDS_TRANSFER))
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = unloadsFIFunds.TransactionCount + mType.TransactionCount;
                                unloadsFIFunds.TransactionCount = value;
                                //calculate Transaction Amount
                                if (unloadsFIFunds.DBCR1.Equals(mType.DBCR1))
                                    value = unloadsFIFunds.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = unloadsFIFunds.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    unloadsFIFunds.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                unloadsFIFunds.TransactionAmount = value;
                                //calculate Fee Amount
                                value = unloadsFIFunds.FeeAmount + mType.FeeAmount;
                                unloadsFIFunds.FeeAmount = value;
                                //calculate Total Amount
                                if (unloadsFIFunds.DBCR2.Equals(mType.DBCR2))
                                    value = unloadsFIFunds.TotalAmount + mType.TotalAmount;
                                else
                                    value = unloadsFIFunds.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    unloadsFIFunds.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                unloadsFIFunds.TotalAmount = value;
                            }
                            //VS4637 Add LOADS BYPASS    
                            else if (mType.Transaction.Equals(TransactionType.TransactionOption.LOADS_BYPASS))
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = loadsByPass.TransactionCount + mType.TransactionCount;
                                loadsByPass.TransactionCount = value;
                                //calculate Transaction Amount
                                if (loadsByPass.DBCR1.Equals(mType.DBCR1))
                                    value = loadsByPass.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = loadsByPass.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    loadsByPass.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                loadsByPass.TransactionAmount = value;
                                //calculate Fee Amount
                                value = loadsByPass.FeeAmount + mType.FeeAmount;
                                loadsByPass.FeeAmount = value;
                                //calculate Total Amount
                                if (loadsByPass.DBCR2.Equals(mType.DBCR2))
                                    value = loadsByPass.TotalAmount + mType.TotalAmount;
                                else
                                    value = loadsByPass.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    loadsByPass.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                loadsByPass.TotalAmount = value;
                            }
                            else if (mType.Transaction.Equals(TransactionType.TransactionOption.RELOADS_BYPASS))
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = reloadsByPass.TransactionCount + mType.TransactionCount;
                                reloadsByPass.TransactionCount = value;
                                //calculate Transaction Amount
                                if (reloadsByPass.DBCR1.Equals(mType.DBCR1))
                                    value = reloadsByPass.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = reloadsByPass.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    reloadsByPass.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                reloadsByPass.TransactionAmount = value;
                                //calculate Fee Amount
                                value = reloadsByPass.FeeAmount + mType.FeeAmount;
                                reloadsByPass.FeeAmount = value;
                                //calculate Total Amount
                                if (reloadsByPass.DBCR2.Equals(mType.DBCR2))
                                    value = reloadsByPass.TotalAmount + mType.TotalAmount;
                                else
                                    value = reloadsByPass.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    reloadsByPass.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                reloadsByPass.TotalAmount = value;
                            }
                            //VS4637
                            else if (mType.Transaction.Equals(TransactionType.TransactionOption.LOADS_MERCHANT_ACQ))
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = loadsMerchACQ.TransactionCount + mType.TransactionCount;
                                loadsMerchACQ.TransactionCount = value;
                                //calculate Transaction Amount
                                if (loadsMerchACQ.DBCR1.Equals(mType.DBCR1))
                                    value = loadsMerchACQ.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = loadsMerchACQ.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    loadsMerchACQ.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                loadsMerchACQ.TransactionAmount = value;
                                //calculate Fee Amount
                                value = loadsMerchACQ.FeeAmount + mType.FeeAmount;
                                loadsMerchACQ.FeeAmount = value;
                                //calculate Total Amount
                                if (loadsMerchACQ.DBCR2.Equals(mType.DBCR2))
                                    value = loadsMerchACQ.TotalAmount + mType.TotalAmount;
                                else
                                    value = loadsMerchACQ.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    loadsMerchACQ.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                loadsMerchACQ.TotalAmount = value;
                            }
                            else if (mType.Transaction.Equals(TransactionType.TransactionOption.RELOADS_MERCHANT_ACQ))
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = reloadsMerchACQ.TransactionCount + mType.TransactionCount;
                                reloadsMerchACQ.TransactionCount = value;
                                //calculate Transaction Amount
                                if (reloadsMerchACQ.DBCR1.Equals(mType.DBCR1))
                                    value = reloadsMerchACQ.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = reloadsMerchACQ.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    reloadsMerchACQ.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                reloadsMerchACQ.TransactionAmount = value;
                                //calculate Fee Amount
                                value = reloadsMerchACQ.FeeAmount + mType.FeeAmount;
                                reloadsMerchACQ.FeeAmount = value;
                                //calculate Total Amount
                                if (reloadsMerchACQ.DBCR2.Equals(mType.DBCR2))
                                    value = reloadsMerchACQ.TotalAmount + mType.TotalAmount;
                                else
                                    value = reloadsMerchACQ.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    reloadsMerchACQ.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                reloadsMerchACQ.TotalAmount = value;
                            }
                            else if (mType.Transaction.Equals(TransactionType.TransactionOption.UNLOADS_MERCHANT_ACQ))
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = unloadsMerchACQ.TransactionCount + mType.TransactionCount;
                                unloadsMerchACQ.TransactionCount = value;
                                //calculate Transaction Amount
                                if (unloadsMerchACQ.DBCR1.Equals(mType.DBCR1))
                                    value = unloadsMerchACQ.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = unloadsMerchACQ.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    unloadsMerchACQ.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                unloadsMerchACQ.TransactionAmount = value;
                                //calculate Fee Amount
                                value = unloadsMerchACQ.FeeAmount + mType.FeeAmount;
                                unloadsMerchACQ.FeeAmount = value;
                                //calculate Total Amount
                                if (unloadsMerchACQ.DBCR2.Equals(mType.DBCR2))
                                    value = unloadsMerchACQ.TotalAmount + mType.TotalAmount;
                                else
                                    value = unloadsMerchACQ.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    unloadsMerchACQ.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                unloadsMerchACQ.TotalAmount = value;
                            }
                            //VS3594 Add UNLOADS BYPASS    
                            else if (mType.Transaction.Equals(TransactionType.TransactionOption.UNLOADS_BYPASS))
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = unloadsByPass.TransactionCount + mType.TransactionCount;
                                unloadsByPass.TransactionCount = value;
                                //calculate Transaction Amount
                                if (unloadsByPass.DBCR1.Equals(mType.DBCR1))
                                    value = unloadsByPass.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = unloadsByPass.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    unloadsByPass.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                unloadsByPass.TransactionAmount = value;
                                //calculate Fee Amount
                                value = unloadsByPass.FeeAmount + mType.FeeAmount;
                                unloadsByPass.FeeAmount = value;
                                //calculate Total Amount
                                if (unloadsByPass.DBCR2.Equals(mType.DBCR2))
                                    value = unloadsByPass.TotalAmount + mType.TotalAmount;
                                else
                                    value = unloadsByPass.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    unloadsByPass.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                unloadsByPass.TotalAmount = value;
                            }
                            else if (mType.Transaction.Equals(TransactionType.TransactionOption.TOTAL_LOAD_UNLOAD_ACTIVITY))
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = totalLoadsUnloads.TransactionCount + mType.TransactionCount;
                                totalLoadsUnloads.TransactionCount = value;
                                //calculate Transaction Amount
                                if (totalLoadsUnloads.DBCR1.Equals(mType.DBCR1))
                                    value = totalLoadsUnloads.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = totalLoadsUnloads.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    totalLoadsUnloads.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                totalLoadsUnloads.TransactionAmount = value;
                                //calculate Fee Amount
                                value = totalLoadsUnloads.FeeAmount + mType.FeeAmount;
                                totalLoadsUnloads.FeeAmount = value;
                                //calculate Total Amount
                                if (totalLoadsUnloads.DBCR2.Equals(mType.DBCR2))
                                    value = totalLoadsUnloads.TotalAmount + mType.TotalAmount;
                                else
                                    value = totalLoadsUnloads.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    totalLoadsUnloads.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                totalLoadsUnloads.TotalAmount = value;
                            }
                            else //must be Manual Adjustment via order by elimination
                            {
                                //calculate Transaction Count
                                value = 0;
                                value = manualAdj.TransactionCount + mType.TransactionCount;
                                manualAdj.TransactionCount = value;
                                //calculate Transaction Amount
                                if (manualAdj.DBCR1.Equals(mType.DBCR1))
                                    value = manualAdj.TransactionAmount + mType.TransactionAmount;
                                else
                                    value = manualAdj.TransactionAmount - mType.TransactionAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    manualAdj.DBCR1 = mType.DBCR1;
                                    value = Math.Abs(value);
                                }
                                manualAdj.TransactionAmount = value;
                                //calculate Fee Amount
                                value = manualAdj.FeeAmount + mType.FeeAmount;
                                manualAdj.FeeAmount = value;
                                //calculate Total Amount
                                if (manualAdj.DBCR2.Equals(mType.DBCR2))
                                    value = manualAdj.TotalAmount + mType.TotalAmount;
                                else
                                    value = manualAdj.TotalAmount - mType.TotalAmount;
                                //check the values to determine if we need to switch DB CR flag..
                                if (value < 0)
                                {
                                    manualAdj.DBCR2 = mType.DBCR2;
                                    value = Math.Abs(value);
                                }
                                manualAdj.TotalAmount = value;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(this, ex.Message + Environment.NewLine + ex.StackTrace, "Error: AddPrepaidCardTotalSummary");
                        }
                    }
                }
            }

            if (_prepaidcardtotalsummary.Count > 0)
            {
                //clear the list before adding the new values...
                CalculatedPrePaidTotalList.Clear();
                CalculatedPrePaidTotalList.Add(loadsFIFunds);
                CalculatedPrePaidTotalList.Add(reloadsFIFunds);
                CalculatedPrePaidTotalList.Add(unloadsFIFunds);
                CalculatedPrePaidTotalList.Add(loadsMerchPOSFunding);
                CalculatedPrePaidTotalList.Add(reloadsByPass);
                CalculatedPrePaidTotalList.Add(unloadsByPass);
                CalculatedPrePaidTotalList.Add(reloadsMerchACQ);
                CalculatedPrePaidTotalList.Add(unloadsMerchACQ);
                CalculatedPrePaidTotalList.Add(totalLoadsUnloads);
                CalculatedPrePaidTotalList.Add(manualAdj);
                bindingSourcePrePaidCardTotalAggregate.DataSource = CalculatedPrePaidTotalList;
                bindingSourcePrePaidCardTotalAggregate.ResetBindings(false);
            }
        }

        /// <summary>
        /// Parse the Global summary data in the VisaDPS report.
        /// </summary>
        /// <param name="data"></param>
        public void ProcessGlobalSummaryData(String data)
        {
            //VS4080 Global Summary Report does not always contain the \r\n at the end of the Net Change row. If this row is
            // the last row on the page, then there will be no \r\n appended at the end of the string and thus we do not parse the data out correctly.
            //Fix: append the \r\n onto the data if it does not exist.
            if(!data.EndsWith("\r\n"))
            {
                data += "\r\n";
            }

            Regex pattern = new Regex("(.*\r\n)");
            MatchCollection mCollection = pattern.Matches(data);
            string rowData;
            if (mCollection.Count > 0)
            {
                mGlobalSummaryData = new GlobalSummaryBranch();
                //VS4583 Changed to Prepaid because now there can be Beken reports also. Removed Map.
                mGlobalSummaryData.ClientID = "Prepaid";
                bool isTransactionType = true;
                for (int i = 0; i < mCollection.Count; i++)
                {
                    rowData = mCollection[i].Value;
                    rowData = rowData.Trim();

                    if(rowData.StartsWith("FUNDS POOL STATUS:")) //we are done processing the data for various different TransactionTypes.
                    {
                        isTransactionType = false;
                        i++; //advance ahead another row because we don't need to deal with the column headers..
                        continue;
                    }

                    //The regex pattern returns every row of data. We do not need data from
                    //row 0 or the first row and any row that starts off with hyphen -
                    //VS4080 we also do not need empty rows either or rows that start with TRANSACTION TYPE
                    if (i == 0 || rowData.StartsWith("-") || rowData.Length == 0 || rowData.StartsWith("TRANSACTION TYPE"))
                        continue;
                    else //we have good data to work with...
                    {
                        if(isTransactionType)
                            ParseTransactionString(rowData);
                        else
                            ParsePoolStatusTransactionString(rowData);

                    }

                    //VS4080
                    if(rowData.StartsWith("NET CHANGE"))
                    {
                        //If we see this data, we are done.
                        break;
                    }


                }

                //Now we need to process the data for the FundsPoolType...

            }
        }

        /// <summary>
        /// Parse the Pool Status information that occurs at the end of a VisaDPS report (In the Global Summary).
        /// </summary>
        /// <param name="inData"></param>
        public void ParsePoolStatusTransactionString(String inData)
        {
            string str = inData;
            FundsPoolStatusType mFundsPoolType = new FundsPoolStatusType();
            Regex pattern = null;
            MatchCollection mCollection = null;

            try
            {
                //get the name Transaction name
                pattern = new Regex("(^[*A-Z/\\s-]+)");
                mCollection = pattern.Matches(str);
                if(mCollection.Count > 0)
                {
                    mFundsPoolType.FundsPoolStatus = mCollection[0].Value;

                //edge case alert -- don't know why but there could be values with a negative sign in it:
                //usually the DR or CR would indicate what to do with the value and you don't see a negative sign.
                //TRANSACTION DISPUTES                1-          39.95 DR        0.00               39.95 DR
                //
                //Some transaction names can have a hyphen in there which looks like a negative sign and thus the
                //regex would grab that value we don't want. To work around this lets clip off the transaction name!
                    str = str.Substring(mCollection[0].Value.Length);
                }

                //get the FUNDS POOL BALANCE,ACCOUNTS REPORTED,OPEN ACCOUNTS,CLOSED ACCOUNTS,VOIDED ACCOUNTS
                pattern = new Regex("(\\s[0-9,.]+\\s|[0-9,.]+)");
                mCollection = pattern.Matches(str);
                if (mCollection.Count > 0)
                {
                    mFundsPoolType.FundsPoolBalance = decimal.Parse(mCollection[0].Value);
                    mFundsPoolType.AccountsReported = decimal.Parse(mCollection[1].Value);
                    mFundsPoolType.OpenAccounts     = decimal.Parse(mCollection[2].Value);
                    mFundsPoolType.ClosedAccounts   = decimal.Parse(mCollection[3].Value);
                    mFundsPoolType.VoidedAccounts   = decimal.Parse(mCollection[4].Value);
                }

                //get the CR|DR values...
                pattern = new Regex("(CR|DR)");
                mCollection = pattern.Matches(str);
                if (mCollection.Count > 0)
                {
                    mFundsPoolType.DBCR = mCollection[0].Value;
                }

            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Global Summary Report Processing Error", MessageBoxButtons.OK);
            }
            mGlobalSummaryData.FundsPoolStatusList.Add(mFundsPoolType);
        }

        /// <summary>
        /// Parse the Global Summary reports different transaction types and their
        /// associated values.
        /// </summary>
        /// <param name="inData"></param>
        public void ParseTransactionString(String inData)
        {
            string str = inData;
            TransactionType aTransaction = new TransactionType();
            Regex pattern = null;
            MatchCollection mCollection = null;

            try
            {
                //get the name Transaction name
                pattern = new Regex("(^[*A-Z/\\s-]+)");
                mCollection = pattern.Matches(str);
                if(mCollection.Count > 0)
                {
                    aTransaction.Transaction = mCollection[0].Value;
                }
                


                //get the Transaction Amount, Fee Amount, and Total Amount 
                pattern = new Regex("(\\s[0-9,.]+\\s|[0-9,.]+)");
                mCollection = pattern.Matches(str);
                if (mCollection.Count > 0)
                {
                    aTransaction.TransactionCount = decimal.Parse(mCollection[0].Value);
                    aTransaction.TransactionAmount = decimal.Parse(mCollection[1].Value);
                    aTransaction.FeeAmount = decimal.Parse(mCollection[2].Value);
                    aTransaction.TotalAmount = decimal.Parse(mCollection[3].Value);
                }

                //get the CR|DR values...

                pattern = new Regex("(CR|DR)");
                mCollection = pattern.Matches(str);
                if (mCollection.Count > 0)
                {
                    //if we have 3 parsed values then we can insert them in order.
                    if (mCollection.Count == 3)
                    {
                        aTransaction.DBCR1 = mCollection[0].Value;
                        aTransaction.DBCR2 = mCollection[1].Value;
                        aTransaction.DBCR3 = mCollection[2].Value;
                    }

                    //if we have 2 then we need to determine where to insert the data
                    if (mCollection.Count == 2)
                    {
                        //if we have no value here then we know that there are no Fees...
                        if (aTransaction.TransactionAmount == 0)
                        {
                            aTransaction.DBCR2 = mCollection[0].Value;
                            aTransaction.DBCR3 = mCollection[1].Value;
                        }
                        else //We have a valid value for TransactionAmount
                        {
                            aTransaction.DBCR1 = mCollection[0].Value;
                            aTransaction.DBCR3 = mCollection[1].Value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Global Summary Report Processing Error", MessageBoxButtons.OK);
            }
            mGlobalSummaryData.TransactionList.Add(aTransaction);
        }

        void _processVisaReport_EndofReport(string data)
        {
            //actually the background worker thread should be done if we get into here... no need to shut it down because it will already be done.
            
            //VS3623  Need to check if bgw thread is still up. If so then turn it off because we are done.
            //if (bgw.IsBusy)
            //{
            //    lock (countLock)
            //    {
            //        if(debug)
            //            System.Diagnostics.Trace.WriteLine("EndofReport encountered ThreadID: " + System.Threading.Thread.CurrentThread.ManagedThreadId);
                    
            //        System.Threading.Monitor.Pulse(countLock);
            //        bgw.CancelAsync();
            //    }
            //}


            //we need to check the BranchInfo.xml file to restore user configured data values for certain ClientBranch.

            LogEvent("Process Visa DPS Report Completed");
            
            MergeClientBranchInfo();

            //sync up the databinding source to the datagrid.
            bindingSourceGPRDetails.DataSource = _gprdetail;
            ultraGridGPRDetails.DataSource = bindingSourceGPRDetails;

            bindingSourceGiftCardDetails.DataSource = _giftdetail;
            ultraGridGiftCardDetails.DataSource = bindingSourceGiftCardDetails;

            bindingSourceGiftSummary.DataSource = _giftsummary;
            ultraGridGiftSummary.DataSource = bindingSourceGiftSummary;

            bindingSourceGPRSummary.DataSource = _gprsummary;
            ultraGridGPRSummary.DataSource = bindingSourceGPRSummary;

            bindingSourcePrepaidCardTotals.DataSource = _prepaidcardtotalsummary;
            ultraGridPrePaidCardTotals.DataSource = bindingSourcePrepaidCardTotals;

            bindingSourceUnrecognizedClientID.DataSource = _unrecognizedid;
            ultraGridUnrecognizedClientID.DataSource = bindingSourceUnrecognizedClientID;

            bindingSourceIgnoreReport.DataSource = _ignorereport;
            ultraGridIgnoreReport.DataSource = bindingSourceIgnoreReport;

            //optimization for addressing issue with the expansion indicator for the Ultradatagrid in which
            //the indicator is default shown even if there are no child band rowns.
            int largestIndex = ultraGridGPRDetails.DisplayLayout.Bands.Count;
            if (ultraGridGiftCardDetails.DisplayLayout.Bands.Count > largestIndex)
                largestIndex = ultraGridGiftCardDetails.DisplayLayout.Bands.Count;
            if (ultraGridGiftSummary.DisplayLayout.Bands.Count > largestIndex)
                largestIndex = ultraGridGiftSummary.DisplayLayout.Bands.Count;
            if (ultraGridGPRSummary.DisplayLayout.Bands.Count > largestIndex)
                largestIndex = ultraGridGPRSummary.DisplayLayout.Bands.Count;
            if (ultraGridPrePaidCardTotals.DisplayLayout.Bands.Count > largestIndex)
                largestIndex = ultraGridPrePaidCardTotals.DisplayLayout.Bands.Count;
            if (ultraGridUnrecognizedClientID.DisplayLayout.Bands.Count > largestIndex)
                largestIndex = ultraGridUnrecognizedClientID.DisplayLayout.Bands.Count;
            if (ultraGridIgnoreReport.DisplayLayout.Bands.Count > largestIndex)
                largestIndex = ultraGridIgnoreReport.DisplayLayout.Bands.Count;

            //set the overriding values for the expansion icons used by the ultradatagrid. That way the default behavior
            //of displaying the expansion icon is not displayed for each row of data even for rows that do not have
            //child rows within them.
            //We don't want to perform multiple for loops for each grid, instead lets
            //go through them in one pass by using the largest index as the overall loop...
            Infragistics.Win.UltraWinGrid.UltraGridBand band;
            try
            {
                for (int i = 0; i < largestIndex; i++)
                {
                    if (i < ultraGridGPRDetails.DisplayLayout.Bands.Count)
                    {
                        band = ultraGridGPRDetails.DisplayLayout.Bands[i];
                        band.Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay;
                    }

                    if (i < ultraGridGiftCardDetails.DisplayLayout.Bands.Count)
                    {
                        band = ultraGridGiftCardDetails.DisplayLayout.Bands[i];
                        band.Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay;
                    }

                    if (i < ultraGridGiftSummary.DisplayLayout.Bands.Count)
                    {
                        band = ultraGridGiftSummary.DisplayLayout.Bands[i];
                        band.Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay;
                    }

                    if (i < ultraGridGPRSummary.DisplayLayout.Bands.Count)
                    {
                        band = ultraGridGPRSummary.DisplayLayout.Bands[i];
                        band.Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay;
                    }

                    if (i < ultraGridPrePaidCardTotals.DisplayLayout.Bands.Count)
                    {
                        band = ultraGridPrePaidCardTotals.DisplayLayout.Bands[i];
                        band.Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay;
                    }

                    if (i < ultraGridUnrecognizedClientID.DisplayLayout.Bands.Count)
                    {
                        band = ultraGridPrePaidCardTotals.DisplayLayout.Bands[i];
                        band.Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay;
                    }

                    if (i < ultraGridIgnoreReport.DisplayLayout.Bands.Count)
                    {
                        band = ultraGridPrePaidCardTotals.DisplayLayout.Bands[i];
                        band.Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay;
                    }
                }
            }
            catch (Exception e)
            {
                TraceLine(e.Message);
            }

            try
            {
                ProcessGlobalSummaryData(_processVisaReport.Data);
                bindingSourceGlobalSummaryBranch.DataSource = mGlobalSummaryData;
                ultraGridGlobalSummary.DataSource = bindingSourceGlobalSummaryBranch;


                //unrecognized id occured. bring up the edit branch info window to peruse..
                if (_unrecognizedid.Count > 0)
                    buttonEdit_Click(this, new EventArgs());
                AddPrepaidCardTotalSummary();
                GenerateExcelPreviewTable();
                enableGUI();
                LogEvent("Finished processing Visa Report");

                //set the globalsummary data into ProcessFinSumReport so that it can access the information to build the CFPRecon table.
                _processFinSumReport.GSummaryBranch = mGlobalSummaryData;

            }
            catch (Exception ex2)
            {
                TraceLine(ex2.Message);
            }
        }

        /// <summary>
        /// _processVisaReport_BranchChanged notification that gets notified when a new client branch is detected
        /// </summary>
        /// <param name="newBranch"></param>
        void _processVisaReport_BranchChanged(ClientBranch newBranch)
        {
            TraceLine("ThreadID BranchChanged: ");
            //add the ClientBranch objects from the _processVisaReport object into the correct bins...
            if (newBranch != null && newBranch.ClientID.Length > 0)
            {
                try
                {
                    switch (_processVisaReport.BranchInfo.Group)
                    {
                        case PrePaidCardsTotal:
                            {
                                //crazy report. SARPage identifier is not a unique value for SubClient ID in the prepaid card total summary.
                                //NORT000 and others have two different SARPages. Go figure? How ever the second SARPage associated with SubClient ID
                                //does not contain any information that we need. It just shows up in the table twice.

                                _prepaidcardtotalsummary.Add(newBranch);
                            }
                            break;

                        case GPRSummary:
                            {
                                _gprsummary.Add(newBranch);

                            }
                            break;

                        case GiftCardSummary:
                            {
   
                                _giftsummary.Add(newBranch);
      
                            }
                            break;

                        case GPRBranchDetail:
                            {
                                _gprdetail.Add(newBranch);
                            }
                            break;

                        case GiftBranchDetail:
                            {
                                _giftdetail.Add(newBranch);
                            }
                            break;

                        case IgnoreReport:
                            {
                                _ignorereport.Add(newBranch);
                            }
                            break;

                        default: //unrecognized...
                            {
                                _unrecognizedid.Add(newBranch);
                            }
                            break;
                    }
                    if (!ClientBranchList.Contains(newBranch))
                        ClientBranchList.Add(newBranch);
                }
                catch (Exception ex)
                {
                    System.Console.WriteLine(ex.Message);
                }
            }
        }

        /// <summary>
        /// Check the running report date and see if it matches up with the current date. If not then we place a warning
        /// message in the Report Date textfield.
        /// </summary>
        /// <param name="date"></param>
        void _processVisaReport_ReportDate(string date)
        {
            DateTime visaReport = DateTime.Parse(date);
            visaReportRunDate = date;
            if(!visaReport.Date.Equals(DateTime.Today.Date))
            {
                if (this.textBoxReportDate.InvokeRequired)
                {
                    TextboxReportDateCallback tcb = new TextboxReportDateCallback(_processVisaReport_ReportDate);
                    this.Invoke(tcb, new object[] { date });
                }
                else
                    textBoxReportDate.Text = "Warning: report's RUNDATE does not match today's date: " + visaReport.ToShortDateString();
            }
        }

        /// <summary>
        /// Notification event when we reach the end of a report.
        /// </summary>
        /// <param name="data"></param>
        void _processFinSumReport_EndofReport()
        {
            //processing of the report has completed.
            //now we need to convert the data retrieved into CFPReport format...
            _processFinSumReport.ConvertDataToCFPRecon();
            enableGUI();
            LogEvent("Finished processing Financial Summary Report");

        }

        /// <summary>
        /// Notification even when we reach the end of a report.
        /// </summary>
        void _processWireReport_EndofReport()
        {
            LogEvent("Finished processing wire confirmation report");
            //export the contents to excel...
            if (mExcel == null)
            {
                mExcel = new MAPacificExportExcel();
                mExcel.ExcelDirectory = UserSettings.Current.ExcelDirectory;
            }
            mExcel.ExportWireConfirmationReport(_processWireReport.UltraGridFinancialSummary, _processWireReport.RunDate);

        }

        #region Background Worker events
   
        
        //VS3623
        /// <summary>
        /// BackgroundWorker thread has completed its task and now we need to reset some data and re-enable the gui interface.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            
            TraceLine("BGW ThreadID workercompleted");
            sw.Stop();
            _processVisaReport.ReportCancel = true;
            //if (e.Cancelled)
            //    LogEvent("Process Report Cancelled");
            //else
            //    LogEvent("Process Report Completed");
            //textBoxReportDate.Text = "";
            enableGUI();
            if (debug)
            {
                MessageBox.Show("Time taken to process the report in ticks: " + sw.ElapsedTicks);
            }
            sw.Reset();

        }

        //VS3623
        /// <summary>
        /// Loop until the background thread has finished processing the VisaDPS report
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            while (!_processVisaReport.ReportCancel && !_processVisaReport.ReportFinish && !bgwVisaDPS.CancellationPending)
            {
                lock (countLock)
                {
                    TraceLine("BGW ThreadID");
                    
                    _processVisaReport.ParseData();


                    if (!_processVisaReport.ReportFinish && !_processVisaReport.ReportCancel) //flag could have been changed after the ParseData method. Check again...
                    {
                        //VS3623 Refactor logic to transmit function key in the background thread. We will transmit the function key and then place the thread in a wait state.
                        //We are waiting for the keyboard unlock message event sent from the OVApi which in turn will signal this thread to process the data.
                //        ovApi.EmitFunctionKey(_processVisaReport.FunctionKey);

                        //swapping out the ovApi.EmitFunctionKey call because sometimes the function key does not get transmitted to OV. Instead use the
                        //CrtTrigger function and check the results. If the function key is sent successfully, the return string value is "OK".
                        string result = "";
                        do
                        {
                            result = ovApi.CrtTrigger("FUNCKEY", _processVisaReport.FunctionKey);
                            TraceLine("CrtTrigger result: " + result);
                        } while (result.Length == 0);
                         
                        TraceLine("Parse Data finished - Transmitted FunctionKey for page: " + _processVisaReport.getSarPage());
                        TraceLine("Monitor Wait called from threadID: ");
   
                        System.Threading.Monitor.Wait(countLock);
                    }
                }
            }

            //if (bgw.CancellationPending)
            //{
            //    e.Cancel = true;
            //}

        }

        void bgwFinancialSummary_DoWork(object sender, DoWorkEventArgs e)
        {
            while (!_processFinSumReport.ReportFinish)
            {
                lock (countLock)
                {
                    TraceLine("BGW ThreadID");

                    _processFinSumReport.ParseData();

                    //VS3623 Refactor logic to transmit function key in the background thread. We will transmit the function key and then place the thread in a wait state.
                    //We are waiting for the keyboard unlock message event sent from the OVApi which in turn will signal this thread to process the data.
                    //        ovApi.EmitFunctionKey(_processVisaReport.FunctionKey);

                    //swapping out the ovApi.EmitFunctionKey call because sometimes the function key does not get transmitted to OV. Instead use the
                    //CrtTrigger function and check the results. If the function key is sent successfully, the return string value is "OK".
                    string result = "";
                    do
                    {
                        result = ovApi.CrtTrigger("FUNCKEY", _processFinSumReport.FunctionKey);
                        TraceLine("CrtTrigger result: " + result);
                    } while (result.Length == 0);

                    TraceLine("Parse Data finished - Transmitted FunctionKey for page: " + _processFinSumReport.SarPage);
                    TraceLine("Monitor Wait called from threadID: ");

                    System.Threading.Monitor.Wait(countLock);
                }
            }

            if (_processFinSumReport.Error)
            {
                TraceLine("Error encountered parsing Financial Summary Report");
                LogEvent("Error processing Financial Summary Report. Are you in the correct report?");
            }
        }

        void bgwFinancialSummary_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            TraceLine("BGW ThreadID workercompleted");
            enableGUI();
        }

        void bgwDPWireConfReport_DoWork(object sender, DoWorkEventArgs e)
        {
            while (!_processWireReport.ReportFinish)
            {
                lock (countLock)
                {
                    TraceLine("BGW ThreadID");

                    _processWireReport.ParseData();

                    if (_processWireReport.ReportFinish) //if after ParseData that we reached the end of a report then break out.
                        break;

                    //VS3623 Refactor logic to transmit function key in the background thread. We will transmit the function key and then place the thread in a wait state.
                    //We are waiting for the keyboard unlock message event sent from the OVApi which in turn will signal this thread to process the data.
                    //        ovApi.EmitFunctionKey(_processVisaReport.FunctionKey);

                    //swapping out the ovApi.EmitFunctionKey call because sometimes the function key does not get transmitted to OV. Instead use the
                    //CrtTrigger function and check the results. If the function key is sent successfully, the return string value is "OK".
                    string result = "";
                    do
                    {
                        result = ovApi.CrtTrigger("FUNCKEY", _processWireReport.FunctionKey);
                        TraceLine("CrtTrigger result: " + result);

                    } while (result.Length == 0);

                    TraceLine("Parse Data finished - Transmitted FunctionKey for page: " + _processWireReport.SarPage);
                    TraceLine("Monitor Wait called from threadID: ");

                    System.Threading.Monitor.Wait(countLock);
                }
            }

            if (_processWireReport.Error)
            {
                TraceLine("Error encountered parsing Wire Confirmation Report");
                LogEvent("Error processing Wire Confirmation Report. Are you in the correct report?");
            }
        }

        void bgwDPWireConfReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            TraceLine("BGW ThreadID workercompleted");
            enableGUI();
        }

        #endregion

        #region form related events

        /// <summary>
        /// Form close handler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReportUtilityMainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void applicationSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SettingsForm sfDialog = new SettingsForm();
            sfDialog.StartPosition = FormStartPosition.CenterParent;
            if (sfDialog.ShowDialog(this) == DialogResult.OK)
            {

            }
            sfDialog.Dispose();
        }

        /// <summary>
        /// ReportUtilityMainForm_Load set the value for the textBoxReportDate with
        /// the current date and label: "PrepaidImport(mm/dd/yyyy)"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReportUtilityMainForm_Load(object sender, EventArgs e)
        {
            textBoxReportDate.Text = "PrepaidImport(" + DateTime.Today.ToShortDateString() + ")";
        }



        /// <summary>
        /// Start the VISADPS report automation so that report contents can
        /// be processed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonProcess_Click(object sender, EventArgs e)
        {
            //VS3623 Reset the TextBoxReportDate to todays date. This value can change via user modification.
            //VS4142 Change the PrepaidImport to just Import because now we need to create excel files that are either PrepaidImport or GiftImport.
            textBoxReportDate.Text = "Import(" + DateTime.Today.ToShortDateString() + ")";

            if (ovApi.CurrentSessionState == OvDotNet.OvSessionState.Disconnected || ovApi.CurrentSessionState == OvDotNet.OvSessionState.FailedLinkage)
            {
                LogEvent("There is no active connected OV 8.1 session");
                return;
            }
            //TestPopulateClientBranch();

            LogEvent("Processing Visa Report");
            _processVisaReport.ResetData();
            ResetDataDefaults(true);
         //   _processVisaReport.ReportFinish = false; --> reset this value in ResetData.
            disableGUI();
            sw.Start();
            //VS3623 Handle the processing of the report asynchronously on the background worker thread.
            bgwVisaDPS.RunWorkerAsync(); 
        }

        /// <summary>
        /// Start processing the Financial Summary Report
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonProcFinSummary_Click(object sender, EventArgs e)
        {
            if (ovApi.CurrentSessionState == OvDotNet.OvSessionState.Disconnected || ovApi.CurrentSessionState == OvDotNet.OvSessionState.FailedLinkage)
            {
                LogEvent("There is no active connected OV 8.1 session");
                return;
            }

            disableGUI();

            LogEvent("Processing Financial Summary Report");
            _processFinSumReport.ResetData();

            bgwFinancialSummary.RunWorkerAsync();
        }

        private void buttonProcessWireConfReport_Click(object sender, EventArgs e)
        {
            if (ovApi.CurrentSessionState == OvDotNet.OvSessionState.Disconnected || ovApi.CurrentSessionState == OvDotNet.OvSessionState.FailedLinkage)
            {
                LogEvent("There is no active connected OV 8.1 session");
                return;
            }

            disableGUI();
            LogEvent("Processing DPS Wire Confirmation Report");
            _processWireReport.ResetData();
            _processWireReport.SetDataContainer(myDataContainer);
            bgwDPWireConfReport.RunWorkerAsync();
        }
        /// <summary>
        /// Launch the Edit preferences panel so useds can modify the Client Branch information
        /// for the listed branches.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonEdit_Click(object sender, EventArgs e)
        {
            if (this.textBoxReportDate.InvokeRequired)
            {
                ShowEditFormCallback sfcb = new ShowEditFormCallback(showEditDialog);
                this.Invoke(sfcb);
            }
            else
                showEditDialog();
        }

        private void showEditDialog()
        {
            SubClientIDPreferencesForm mForm = new SubClientIDPreferencesForm(ClientBranchList);
            mForm.StartPosition = FormStartPosition.CenterParent;
            if (mForm.ShowDialog(this) == DialogResult.OK)
            {
                UpdateDataTables(mForm.ReportClientBranchList);
                ClientBranchList = mForm.ReportClientBranchList;
                //VS4229 retrieve the PRCNO infor from the xml file and update that into the  processWireReport object.
                myDataContainer = mForm.MyDContainer;
                if (myDataContainer.WireConfTableList.Count > 0)
                {
                    _processWireReport.SetDataContainer(myDataContainer);
                }
                if (_prepaidcardtotalsummary.Count > 0)
                    GenerateExcelPreviewTable();
            }
            mForm.Dispose();
        }

        private void buttonGenerate_Click(object sender, EventArgs e)
        {
            //retrieve the directory and compose the filename
            try
            {
                //VS4142 Use a bool flag determine report type so that we can set the filename correctly.
                bool isGPRBIN = false;
                bool isBeken = false;
                if (ultraGridExcel.Rows.Count > 0 && ultraGridGlobalSummary.Rows.Count > 0)
                {
                    //VS3611
                    //clone the data from mExcel.ToExcelList so that when we
                    //export the ulgraGridExcel to a Excel file and modify the contents of some cells
                    //during the export event the modified data doesn't wipe out the original data
                    //within the mExcel.ToExcelList.
                    ExportExcelList = new List<ClientBranch>();
                    foreach (ClientBranch sourceBranch in mExcel.ToExcelList)
                    {
                        ExportExcelList.Add(sourceBranch.Clone());
                        if (sourceBranch.BIN.Equals(MAP_GPR_BIN) || (sourceBranch.BIN.Equals(BEKEN_GPR_BIN)))
                            isGPRBIN = true;

                        if (sourceBranch.BIN.Equals(BEKEN_GIFT_BIN) || (sourceBranch.BIN.Equals(BEKEN_GPR_BIN)))
                            isBeken = true;
                    }

                    string fileName = "";
                    if(isGPRBIN)
                    {
                        fileName = "Prepaid" + textBoxReportDate.Text;
                    }
                    else
                    {
                        fileName = "Gift" + textBoxReportDate.Text;
                    }

                    if (isBeken)
                    {
                        fileName = "Beken" + fileName;
                    }

                    bindingSourceClientBranch.DataSource = ExportExcelList;
                    ultraGridExcel.DataSource = bindingSourceClientBranch;
                    //in case user decided to change the excel output directory...
                    mExcel.ExcelDirectory = UserSettings.Current.ExcelDirectory;

                    //Setup and process the data that will be used by the UltraGrid to export out to a worksheet. 
                    ExportExcelTransactionAmount mExportExcelTAmount = new ExportExcelTransactionAmount(_prepaidcardtotalsummary, isBeken);

                    //Setup and process the data that will be used by the UltraGrid to export out to a worksheet. 
                    ExportExcelTransactionCount mExportExcelTCount = new ExportExcelTransactionCount(_prepaidcardtotalsummary, isBeken);

                    //VS4731
                    ExportExcelCardActivityAmount mExportExcelCardAmount = new ExportExcelCardActivityAmount(_prepaidcardtotalsummary, isBeken);
                    //VS4732
                    ExportExcelCardActivityCount mExportExcelCardCount = new ExportExcelCardActivityCount(_prepaidcardtotalsummary, isBeken);
                    
                    mFinancialSummaryReportGrid = _processFinSumReport.UltraGridFinancialSummary;

                    mExcel.ExportToExcel(ultraGridExcel, ultraGridGlobalSummary, mExportExcelTAmount.UltgragridTAmount, mExportExcelTCount.UltgragridTCount, mFinancialSummaryReportGrid, fileName, mExportExcelCardAmount.UltgragridCardAmount, mExportExcelCardCount.UltgragridCardCount);
                    //reset the binding source for the ultraGridExcel back to the clientbranch...
                    bindingSourceClientBranch.DataSource = mExcel.ToExcelList;
                    ultraGridExcel.DataSource = bindingSourceClientBranch;                   
                    LogEvent("Finished exporting report to Excel File");
                }
                //VS4183 Export to excel file the CFPREcon information and the Global Summary Info
                else if (ultraGridExcel.Rows.Count == 0 && ultraGridGlobalSummary.Rows.Count > 0)
                {
                    //VS3611
                    //clone the data from mExcel.ToExcelList so that when we
                    //export the ulgraGridExcel to a Excel file and modify the contents of some cells
                    //during the export event the modified data doesn't wipe out the original data
                    //within the mExcel.ToExcelList.
                    ExportExcelList = new List<ClientBranch>();
                    foreach (ClientBranch sourceBranch in mExcel.ToExcelList)
                    {
                        ExportExcelList.Add(sourceBranch.Clone());
                        if (sourceBranch.BIN.Equals(MAP_GPR_BIN))
                            isGPRBIN = true;
                    }

                    string fileName = "";
                    //VS4583 JIt is possible to have reports that have no client branch in ExcelList. When that happens,
                    //there are no PrepaidImport, Transaction_Account, Transaction_Amount worksheets in the report. MAPacific still wants to see 
                    //a PrepaidImport report so we have to figure out the correct report type running for this edge case.
                    if (isGPRBIN || _processVisaReport.ReportType.Contains("GPR"))
                    {
                        fileName = "Prepaid" + textBoxReportDate.Text;
                    }
                    else
                    {
                        fileName = "Gift" + textBoxReportDate.Text;
                    }

                    //VS4583 need to determine if report is Beken or Map so that we can create the correct file name.
                    if (_processVisaReport.ReportType.Contains("Beken"))
                    {
                        fileName = "Beken" + fileName;
                    }


                    bindingSourceClientBranch.DataSource = ExportExcelList;
                    ultraGridExcel.DataSource = bindingSourceClientBranch;
                    //in case user decided to change the excel output directory...
                    mExcel.ExcelDirectory = UserSettings.Current.ExcelDirectory;

                    mFinancialSummaryReportGrid = _processFinSumReport.UltraGridFinancialSummary;

                    mExcel.ExportToExcel(ultraGridGlobalSummary, mFinancialSummaryReportGrid, fileName);
                    //reset the binding source for the ultraGridExcel back to the clientbranch...
                    bindingSourceClientBranch.DataSource = mExcel.ToExcelList;
                    ultraGridExcel.DataSource = bindingSourceClientBranch;
                    LogEvent("Finished exporting report to Excel File");
                }
                else
                {
                    MessageBox.Show("There are no report data to export.", "Unable to Create Excel File", MessageBoxButtons.OK);
                    LogEvent("Error exporting report to Excel File");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please check the file name and path values again." + Environment.NewLine + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace, "Unable to Create Excel File", MessageBoxButtons.OK);
                LogEvent("Error exporting report to Excel File");
                LogEvent(ex.StackTrace);
            }

        }



        /// <summary>
        /// Generate Individual client reports from the master report parsed from processing the VisaDPS
        /// report.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonGenerateClientReports_Click(object sender, EventArgs e)
        {
            disableGUI();
            //VS4583 Added Beken reports. Need to be able
            //to determine if report is for MAP or Beken.
            bool isMap = false;

            //retrieve the directory and compose the filename
            try
            {
                if (ultraGridExcel.Rows.Count > 0)
                {
                    //VS4142 generate client names that have the correct prefix.
                    //VS4583 Enhance report to include Beken reports
                    string fileName = "";
                    if (mExcel.ToExcelList.Count > 0)
                    {
                        if (mExcel.ToExcelList[0].BIN.Equals(MAP_GPR_BIN))
                        {
                            fileName = "Prepaid " + visaReportRunDate;
                            isMap = true;
                        }
                        else if (mExcel.ToExcelList[0].BIN.Equals(MAP_GIFT_BIN))
                        {
                            fileName = "Gift " + visaReportRunDate;
                            isMap = true;
                        }
                        else if (mExcel.ToExcelList[0].BIN.Equals(BEKEN_GPR_BIN))
                        {
                            fileName = "Prepaid " + "Beken" + visaReportRunDate;
                            isMap = false;
                        }
                        else
                        {
                            fileName = "Gift " + "Beken" + visaReportRunDate;
                            isMap = false;
                        }
                    }
                    else
                        fileName = visaReportRunDate;
                    //in case user decided to change the excel output directory...
                    LogEvent("Starting generate client excel reports...");
                    mExcel.ExcelDirectory = UserSettings.Current.ExcelDirectory;
                    mExcel.GenerateClientExcelReport(ultraGridExcel, fileName, isMap);
                    //reset the binding source for the ultraGridExcel back to the clientbranch...
                    bindingSourceClientBranch.DataSource = mExcel.ToExcelList;
                    ultraGridExcel.DataSource = bindingSourceClientBranch;
                    LogEvent("Finished Generating Client Excel Reports");
                }
                else
                {
                    MessageBox.Show("There are no report data to export.", "Unable to Create Individual Client Excel Files", MessageBoxButtons.OK);
                    LogEvent("Error generating Client Excel Files. Did you process the VisaDPS report first?");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + "Please check the file name and path values again.", "Unable to Create Excel File", MessageBoxButtons.OK);
                LogEvent("Error generating Client Excel Files");
                bindingSourceClientBranch.DataSource = ExportExcelList;
                ultraGridExcel.DataSource = bindingSourceClientBranch;
            }
            enableGUI();
        }

        /// <summary>
        /// Create email drafts and attach the PDF version of the client excel reports.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_emailDraft_Click(object sender, EventArgs e)
        {
            if (mExcel == null || mExcel.isEmailDraftListEmpty())
            {
                MessageBox.Show("There are no client excel reports to create email draft.", "Unable to Create Email Drafts", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                LogEvent("There are no client excel reports to create email draft.");
            }
            else
            {
                LogEvent("Saving email drafts...");
                mExcel.SaveEmailDraft();
                LogEvent("Finished creating the email drafts.");
            }
        }

        //VS3623 Refactor code into the bgw thread's work complete method to activate the user gui and reset some default values.
        //VS3603 placing the parsing logic of the report into background thread.
        //By placing the parsing logic onto a background thread, users can now cancel a processing report.
        /// <summary>
        /// Cancel the report.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonCancelReport_Click(object sender, EventArgs e)
        {
            //VS3623 Handle Cancel action...
            if (bgwVisaDPS.IsBusy)
            {
                LogEvent("Process Visa DPS Report Canceled");
                bgwVisaDPS.CancelAsync();
            }

            if (bgwDPWireConfReport.IsBusy)
            {
                LogEvent("Wire Confirmation Report Canceled");
                bgwDPWireConfReport.CancelAsync();
            }

            if (bgwFinancialSummary.IsBusy)
            {
                LogEvent("Financial Summary Report Canceled");
                bgwFinancialSummary.CancelAsync();
            }
            enableGUI();
        }
        #endregion

        /// <summary>
        /// Generate the preview table that contains the excel data which will be used to export out to the excel file.
        /// </summary>
        private void GenerateExcelPreviewTable()
        {
            if (mExcel == null)
            {
                mExcel = new MAPacificExportExcel(_prepaidcardtotalsummary, _gprsummary, _giftsummary, _gprdetail, _giftdetail, ClientBranchList, _processVisaReport.Data);
            }
            else
            {
                mExcel.SetBindingLists(_prepaidcardtotalsummary, _gprsummary, _giftsummary, _gprdetail, _giftdetail, ClientBranchList, _processVisaReport.Data);
            }
            try
            {
                mExcel.ExcelDirectory = UserSettings.Current.ExcelDirectory;

                if (mExcel.ExcelDirectory.Length == 0)
                {
                    //no directory has been set so lets default to C:\ location
                    mExcel.ExcelDirectory = "C:\\";
                }
            }
            catch (Exception ex)
            {
                //default the location to user profile area instead...
                mExcel.ExcelDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"MAPReportUtility");
                UserSettings.Current.ExcelDirectory = mExcel.ExcelDirectory;
                UserSettings.Save();

            }

            //create the filename for the excel...
            //value in textBoxReportDate is default to current date: PrepaidImport(12/8/2010)
            //However users can type whatever value they want and also if current date does not match the runtime date then a warning
            //message is appended into that textbox.

            string fileName = textBoxReportDate.Text;
            Regex pattern = new Regex("([0-9]+/[[0-9]/[0-9]+)");
            MatchCollection mCollection = pattern.Matches(fileName);
            if (mCollection.Count > 0)
            {
                int dateLength = mCollection[0].Value.Length;
                //we have a match for the date pattern
                //now replace the / values with _ (10/13/2010 --> 10_13_2010
                pattern = new Regex("([0-9]+)");
                mCollection = pattern.Matches(mCollection[0].Value);
                string date = "";
                for (int i = 0; i < mCollection.Count; i++)
                {
                    date += mCollection[i].Value;
                    if (i != mCollection.Count - 1)
                        date += "_";
                }

                //filename matches the default format...
                int index = fileName.IndexOf(mCollection[0].Value);
                if (index > 0 && fileName.StartsWith("PrepaidImport"))
                {
                    fileName = fileName.Substring(0, index);
                    fileName += date + ")";
                    mExcel.ExcelFileName = fileName;
                }
                else
                {
                    //date happens to sit in the middle of the file name...
                    //"FileName 12/25/2010 Quarter4"
                    if (index > 0)
                    {
                        fileName = textBoxReportDate.Text.Substring(0, index);
                        fileName += date;
                        fileName += textBoxReportDate.Text.Substring(index + dateLength);
                        mExcel.ExcelFileName = fileName;
                    }
                }
            }
            else
            {
                mExcel.ExcelFileName = textBoxReportDate.Text;
            }

            mExcel.SortDataIntoMAPFormat();
            ExportExcelList = mExcel.ToExcelList;
            bindingSourceClientBranch.DataSource = ExportExcelList;
            ultraGridExcel.DataSource = bindingSourceClientBranch;
        }

        /// <summary>
        /// Before displaying the parsed Client Branch data into the datagrid table, we need to read in the BranchInfo.xml file of all the
        /// ClientBranch objects.  Then copy over their overriding user specified data into local ClientBranch objects then display them
        /// in the datagrid table.
        /// </summary>
        /// <param name="inList">BindingList of all the ClientBranch data including ones that have updated data recovered from BranchInfo.xml</param>
        private void UpdateDataTables(BindingList<ClientBranch> inList)
        {
            ResetDataDefaults(false);
            //need to add logic to update ClientBranch entries in the various different data tables because user may have updated/changed the category values for some of them.
            foreach (ClientBranch mBranch in inList)
            {
                if (mBranch.Category.Equals(ClientBranch.ClientBranch_Category.GIFT_DETAIL))
                {
                    _giftdetail.Add(mBranch);
                }
                else if (mBranch.Category.Equals(ClientBranch.ClientBranch_Category.GIFT_SUMMARY))
                {
                    _giftsummary.Add(mBranch);
                }
                else if (mBranch.Category.Equals(ClientBranch.ClientBranch_Category.GPR_DETAIL))
                {
                    _gprdetail.Add(mBranch);
                }
                else if (mBranch.Category.Equals(ClientBranch.ClientBranch_Category.GPR_SUMMARY))
                {
                    _gprsummary.Add(mBranch);
                }
                else if (mBranch.Category.Equals(ClientBranch.ClientBranch_Category.IGNORE))
                {
                    _ignorereport.Add(mBranch);
                }
                else if (mBranch.Category.Equals(ClientBranch.ClientBranch_Category.PREPAID_TOTAL_SUMMARY))
                {
                    _prepaidcardtotalsummary.Add(mBranch);
                }
            }

            //update the PrePaidCard Total Summary
            AddPrepaidCardTotalSummary();
        }

        /// <summary>
        /// Reset the various List objects and clear out old data from bindingsources and bingindlists.
        /// </summary>
        /// <param name="resetAll"></param>
        private void ResetDataDefaults(bool resetAll)
        {
            //reset the various different BindingList
            if (!resetAll)
            {
                _prepaidcardtotalsummary.Clear();
                _gprsummary.Clear();
                _giftsummary.Clear();
                _gprdetail.Clear();
                _giftdetail.Clear();
                _ignorereport.Clear();
                _unrecognizedid.Clear();
                return;
            }

            ExportExcelList.Clear();
            bindingSourceClientBranch.DataSource = ExportExcelList;
            bindingSourceClientBranch.ResetBindings(false);
            CalculatedPrePaidTotalList.Clear();
            bindingSourcePrePaidCardTotalAggregate.DataSource = CalculatedPrePaidTotalList;
            bindingSourcePrePaidCardTotalAggregate.ResetBindings(false);
            mGlobalSummaryData.FundsPoolStatusList.Clear();
            mGlobalSummaryData.TransactionList.Clear();
            bindingSourceGlobalSummaryBranch.DataSource = mGlobalSummaryData;
            bindingSourceGlobalSummaryBranch.ResetBindings(false);

            _prepaidcardtotalsummary.Clear();
            _gprsummary.Clear();
            _giftsummary.Clear();
            _gprdetail.Clear();
            _giftdetail.Clear();
            _ignorereport.Clear();
            _unrecognizedid.Clear();
            ClientBranchList.Clear(); //reset the branchlist so when running the report again it will generate updated pre paid total summary
        }

        //$d02
        private void TraceLine(string message)
        {
            if(debug)
                System.Diagnostics.Trace.WriteLine(message + " -- thread ID: " + System.Threading.Thread.CurrentThread.ManagedThreadId);
        }

        #region test methods
        private void TestPopulateClientBranch()
        {
            _prepaidcardtotalsummary = new BindingList<ClientBranch>();
            _gprsummary = new BindingList<ClientBranch>();
            _giftsummary = new BindingList<ClientBranch>();
            _gprdetail = new BindingList<ClientBranch>();
            _giftdetail = new BindingList<ClientBranch>();
            _ignorereport = new BindingList<ClientBranch>();
            _unrecognizedid = new BindingList<ClientBranch>();

            ClientBranch mTest = new ClientBranch();
            mTest.ClientID = "MAPBD000";
            mTest.Name = "MAPBusDev";
            mTest.ExcelSummary = "ALL";
            mTest.Category = ClientBranch.ClientBranch_Category.PREPAID_TOTAL_SUMMARY;
            TransactionType mTestTranType = new TransactionType();
            mTestTranType.FeeAmount = 0.00m;
            mTestTranType.TotalAmount = 275.00m;
            mTestTranType.Transaction = TransactionType.TransactionOption.RELOADS_BYPASS;
            mTestTranType.TransactionAmount = 275.00m;
            mTestTranType.TransactionCount = 1;
            mTestTranType.DBCR1 = "DR";
            mTestTranType.DBCR2 = "DR";
            mTest.TransactionList.Add(mTestTranType);

            mTestTranType = new TransactionType();
            mTestTranType.FeeAmount = 0.00m;
            mTestTranType.TotalAmount = 275.00m;
            mTestTranType.Transaction = TransactionType.TransactionOption.TOTAL_LOAD_UNLOAD_ACTIVITY;
            mTestTranType.TransactionAmount = 275.00m;
            mTestTranType.TransactionCount = 1;
            mTestTranType.DBCR1 = "DR";
            mTestTranType.DBCR2 = "DR";
            mTest.TransactionList.Add(mTestTranType);

            //mTestTranType = new TransactionType();
            //mTestTranType.FeeAmount = 0.00m;
            //mTestTranType.TotalAmount = 150.00m;
            //mTestTranType.Transaction = TransactionType.TransactionOption.TOTAL_LOAD_UNLOAD_ACTIVITY;
            //mTestTranType.TransactionAmount = 150.00m;
            //mTestTranType.TransactionCount = 2;
            //mTestTranType.DBCR1 = "CR";
            //mTestTranType.DBCR2 = "CR";
            //mTest.TransactionList.Add(mTestTranType);
            //_gprdetail.Add(mTest);
            _prepaidcardtotalsummary.Add(mTest);

            mTest = new ClientBranch();
            mTest.ClientID = "MAPBD001";
            mTest.Name = "MAPBusDev1";
            mTest.ExcelSummary = "";
            mTest.Category = ClientBranch.ClientBranch_Category.GPR_DETAIL;
            mTestTranType = new TransactionType();
            mTestTranType.FeeAmount = 0.00m;
            mTestTranType.TotalAmount = 275.00m;
            mTestTranType.Transaction = TransactionType.TransactionOption.RELOADS_BYPASS;
            mTestTranType.TransactionAmount = 275.00m;
            mTestTranType.TransactionCount = 1;
            mTestTranType.DBCR1 = "DR";
            mTestTranType.DBCR2 = "DR";
            mTest.TransactionList.Add(mTestTranType);
            mTestTranType = new TransactionType();
            mTestTranType.FeeAmount = 0.00m;
            mTestTranType.TotalAmount = 275.00m;
            mTestTranType.Transaction = TransactionType.TransactionOption.TOTAL_LOAD_UNLOAD_ACTIVITY;
            mTestTranType.TransactionAmount = 275.00m;
            mTestTranType.TransactionCount = 1;
            mTestTranType.DBCR1 = "DR";
            mTestTranType.DBCR2 = "DR";
            mTest.TransactionList.Add(mTestTranType);
            _gprdetail.Add(mTest);

            bindingSourceGPRDetails.DataSource = _gprdetail;
            ultraGridGPRDetails.DataSource = bindingSourceGPRDetails;
       //     ultraGrid1.DataSource = _gprdetail;
            _processVisaReport_EndofReport("No summary"); 
            
        }
        #endregion

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 aBox = new AboutBox1();
            if (aBox.ShowDialog() == System.Windows.Forms.DialogResult.No)
                aBox.Close();
        }
    }
}

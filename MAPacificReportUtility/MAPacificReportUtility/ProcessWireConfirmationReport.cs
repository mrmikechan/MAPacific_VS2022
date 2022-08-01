using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Infragistics.Win.UltraWinGrid;
using Infragistics.Win;

/*
 * 
 * $d03  03252016    The PRCNo was being parsed with the following pattern:  PRC[0-9]{3} originally consisted of PRC and 3 digits
 *                   Pattern update -- This needs to change from the 3 digits to be both alphanumeric letters so that it can recognize PRCZT1
 * 
 */

namespace MAPacificReportUtility
{
    class ProcessWireConfirmationReport : ProcessReport
    {      
        public ProcessWireConfirmationReport(OvDotNet.OvDotNetApi inApi)
        {
            ovApi = inApi;
            ultragridFinancialSummary = new Infragistics.Win.UltraWinGrid.UltraGrid();

            FontData fd1 = ultragridFinancialSummary.DisplayLayout.Appearance.FontData;
            fd1.Name = "Arial";
            ultragridFinancialSummary.InitializeLayout += new InitializeLayoutEventHandler(ultragridFinancialSummary_InitializeLayout);
            ultragridFinancialSummary.InitializeRow += new InitializeRowEventHandler(ultragridFinancialSummary_InitializeRow);
            processWireItemList = new List <ProcessWireConfirmationItem>();
            mCurrentState = WireConfReportState.ReportSelectionListState;
            mPriorPRCNO = "";
            wireConfTable = new Dictionary<string, bool>();
            specialPRCNOTable = new Dictionary<string, bool>();
            specialPRCNOTable.Add("PRC451", true);
            mKickOutToReportSelectionList = false;

            bRunDateSet = false;
        }

#region properties

        private OvDotNet.OvDotNetApi ovApi;
        private string apiText = "";
        //flag to turn on/off Trace statements
        bool debug = false;

        //VS4229 Add support to save Credit Union to file so that we can read the list of CU
        //to determine which CU have the Wire Confirmation Total column filled with data.
        private DataContainer myContainer;

        //variable to calculate the running column total values
        private decimal mTotalVisaDPSFinancialTrans = 0;
        private decimal mTotalExceptionTrans = 0;
        private decimal mTotalMiscFees = 0;
        private decimal mTotalVisaATMREIMBFees = 0;
        private decimal mTotalVISAATMISAFees = 0;
        private decimal mTotalInterlinkREIMBFees = 0;
        private decimal mTotalInterlinkISAFees = 0;
        private decimal mTotalGrandTotal = 0;
        private decimal mTotalWireConfirmTotal = 0;

        private const string CFP_PRC_NO = "PRC NO.";
        private const string CFP_REF = "REF.";
        private const string CFP_CREDIT_UNION = "CREDIT UNION";

        private const string CFP_VISA_DPS_FINANCIAL_TRANS = "VISA DPS FINANCIAL TRANS";
        private const string CFP_EXCEPTION_TRANS = "EXCEPTION TRANSACTIONS";
        private const string CFP_MISCELLANEOUS_FEES = "MISCELLANEOUS FEES";
        private const string CFP_VISA_ATM_REIMB_FEES = "VISA/ATM REIMB FEES";
        private const string CFP_VISA_ATM_ISA_FEES = "VISA/ATM ISA FEES";
        private const string CFP_INTERLINK_REIMB_FEES = "INTERLINK REIMB FEES";
        private const string CFP_INTERLINK_ISA_FEES = "INTERLINK ISA FEES";
        private const string CFP_GRAND_TOTAL = "GRAND TOTAL";
        private const string CFP_WIRE_CONF_TOTAL = "WIRE CONFIRMATION TOTAL";


        private const string CFP_EF500210_RP03 = "EF500210-RP03";
        private const string CFP_DD0714_D02 = "DD0714-D02";
        private const string CFP_DD0716_D02 = "DD0716-D02";
        private const string CFP_VSS_110_VS = "VSS-110 -VS";
        private const string CFP_VSS_140_VS = "VSS-140 -VS";
        private const string CFP_VSS_110_IL = "VSS-110 -IL";
        private const string CFP_VSS_140_IL = "VSS-140 -IL";

        //constants used for ultragrid access into tables... mainly readability to figure out which columns.
        //private const int COL_PRC_NO = 0;
        //private const int COL_REF = 1;
        //private const int COL_CREDIT_UNION = 2;
        //private const int COL_VISA_FIN_TRAN = 3;
        //private const int COL_EXCEPTION_TRAN = 4;
        //private const int COL_MISC_FEES = 5;
        //private const int COL_VISA_REIMB_FEES = 6;
        //private const int COL_VISA_ISA_FEES = 7;
        //private const int COL_INTERLINK_REIMB_FEES = 8;
        //private const int COL_INTERLINKE_ISA_FEES = 9;
        //private const int COL_GRAND_TOTAL = 10;
        //private const int COL_WIRE_CONF_TOTAL = 11;

        private const string F8Key = "PF8";      //go to next page in Financial Summary Report
        private const string EnterKey = "Enter"; //Entkey to submit the option selected
        private const string F3Key = "PF3";      //kick back out from Financial Summary Report back to Page Index Selection and then back to Report Selection List.


        //bool variable to let us know that we are going back to the Report Selection List screen from Financial Summary Report screen.
        private bool mKickOutToReportSelectionList;
        //Different states that we can be in for the report.
        /// <summary>
        /// Different possible states the report can be in.
        /// </summary>
        public enum WireConfReportState
        {
            ReportSelectionListState = 0,  //At Report Selection List page
            PageIndexSelectionListState,   //At Page Index Selection List page
            SarReportState,                //At a report with SAR page
            FinishState,                   //All sub reports are completed 
        }

        //We only need to run the logic to set the rundate once.
        bool bRunDateSet;

        //Rundate for the report.
        private string mRunDate = "";
        public string RunDate
        {
            get { return mRunDate; }
            set
            {
                mRunDate = string.IsNullOrEmpty(value) ? "" : value.Trim();
            }
        }

        /// <summary>
        /// Determine what the previous report was so we dont select it again from the Report Selection List.
        /// </summary>
        private string mPriorPRCNO = "";
        public string PriorPRCNO
        {
            get { return mPriorPRCNO; }
            set
            {
                if (value != null)
                {
                    mPriorPRCNO = string.IsNullOrEmpty(value) ? "" : value.Trim();
                }
            }
        }

        /// <summary>
        /// Current report state that we are in.
        /// </summary>
        private WireConfReportState mCurrentState;
        public WireConfReportState CurrentState
        {
            get { return mCurrentState; }
            set { mCurrentState = value; }
        }

        
        private bool mError = false;
        public bool Error
        {
            get { return mError; }
        }

        private bool _reportfinish = true; //start off with true because the report isn't running when application is launched.
        public bool ReportFinish
        {
            get { return _reportfinish; }
            set
            {
                _reportfinish = value;
            }
        }

        private int subPage = 0;
        private int _sarpage;
        /// <summary>
        /// </summary>
        public int SarPage
        {
            get { return _sarpage; }
            set
            {
                _sarpage = value;
            }
        }

        private List<ProcessWireConfirmationItem> processWireItemList;
        private Dictionary<string, bool> wireConfTable;

        //special PRCNO Dictionary is used to hold the list of PRCNO that needs to break down the Credit Union name into a replacement name.
        //currently there is only 1 PRCNO (PRC451) that falls into this category.
        private Dictionary<string, bool> specialPRCNOTable;

        private Infragistics.Win.UltraWinGrid.UltraGrid ultragridFinancialSummary;
        public Infragistics.Win.UltraWinGrid.UltraGrid UltraGridFinancialSummary
        {
            get { return ultragridFinancialSummary; }
        }

#endregion

#region ultragrid events

        void ultragridFinancialSummary_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            //Highlight the Total row in the grid so that when exported out to Excel that row is also highlighted.
            if (e.Row.Cells["PRC_NO"].Value.ToString().Length <= 0)
            {
                //find the first occuring empty rown and we should be at the running total row index.
                e.Row.Cells["PRC_NO"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["PRC_NO"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["PRC_NO"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["PRC_NO"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["REF"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["REF"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["REF"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["REF"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["CREDIT_UNION"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["CREDIT_UNION"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["CREDIT_UNION"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["CREDIT_UNION"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["VISA_DPS_FINANCIAL_TRANS"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["VISA_DPS_FINANCIAL_TRANS"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["VISA_DPS_FINANCIAL_TRANS"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["VISA_DPS_FINANCIAL_TRANS"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["EXCEPTION_TRANS"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["EXCEPTION_TRANS"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["EXCEPTION_TRANS"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["EXCEPTION_TRANS"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["MISCELLANEOUS_FEES"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["MISCELLANEOUS_FEES"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["MISCELLANEOUS_FEES"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["MISCELLANEOUS_FEES"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["VISA_ATM_REIMB_FEES"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["VISA_ATM_REIMB_FEES"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["VISA_ATM_REIMB_FEES"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["VISA_ATM_REIMB_FEES"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["VISA_ATM_ISA_FEES"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["VISA_ATM_ISA_FEES"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["VISA_ATM_ISA_FEES"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["VISA_ATM_ISA_FEES"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["INTERLINK_REIMB_FEES"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["INTERLINK_REIMB_FEES"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["INTERLINK_REIMB_FEES"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["INTERLINK_REIMB_FEES"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["INTERLINK_ISA_FEES"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["INTERLINK_ISA_FEES"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["INTERLINK_ISA_FEES"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["INTERLINK_ISA_FEES"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["GRAND_TOTAL"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["GRAND_TOTAL"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["GRAND_TOTAL"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["GRAND_TOTAL"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;

                e.Row.Cells["WIRE_CONFIRMATION_TOTAL"].Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                e.Row.Cells["WIRE_CONFIRMATION_TOTAL"].Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                e.Row.Cells["WIRE_CONFIRMATION_TOTAL"].Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells["WIRE_CONFIRMATION_TOTAL"].Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;


            }
        }

        /// <summary>
        /// Event handler to allow the grid exported to excel to display multiline column data in a cell.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultragridFinancialSummary_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;
            e.Layout.Override.RowSizing = RowSizing.AutoFree;
            e.Layout.Override.CellMultiLine = DefaultableBoolean.True;
            e.Layout.Override.WrapHeaderText = DefaultableBoolean.True;

            e.Layout.Bands[0].Columns["PRC_NO"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["PRC_NO"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["PRC_NO"].Header.Caption = CFP_PRC_NO;

            e.Layout.Bands[0].Columns["REF"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["REF"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["REF"].Header.Caption = CFP_REF;

            e.Layout.Bands[0].Columns["CREDIT_UNION"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["CREDIT_UNION"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["CREDIT_UNION"].Header.Caption = CFP_CREDIT_UNION;

            e.Layout.Bands[0].Columns["VISA_DPS_FINANCIAL_TRANS"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["VISA_DPS_FINANCIAL_TRANS"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["VISA_DPS_FINANCIAL_TRANS"].Header.Caption = CFP_EF500210_RP03;


            e.Layout.Bands[0].Columns["EXCEPTION_TRANS"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["EXCEPTION_TRANS"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["EXCEPTION_TRANS"].Header.Caption = CFP_DD0714_D02;

            e.Layout.Bands[0].Columns["MISCELLANEOUS_FEES"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["MISCELLANEOUS_FEES"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["MISCELLANEOUS_FEES"].Header.Caption = CFP_DD0716_D02;

            e.Layout.Bands[0].Columns["VISA_ATM_REIMB_FEES"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["VISA_ATM_REIMB_FEES"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["VISA_ATM_REIMB_FEES"].Header.Caption = CFP_VSS_110_VS;

            e.Layout.Bands[0].Columns["VISA_ATM_ISA_FEES"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["VISA_ATM_ISA_FEES"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["VISA_ATM_ISA_FEES"].Header.Caption = CFP_VSS_140_VS;

            e.Layout.Bands[0].Columns["INTERLINK_REIMB_FEES"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["INTERLINK_REIMB_FEES"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["INTERLINK_REIMB_FEES"].Header.Caption = CFP_VSS_110_IL;

            e.Layout.Bands[0].Columns["INTERLINK_ISA_FEES"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["INTERLINK_ISA_FEES"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["INTERLINK_ISA_FEES"].Header.Caption = CFP_VSS_140_IL;

            e.Layout.Bands[0].Columns["GRAND_TOTAL"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["GRAND_TOTAL"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["GRAND_TOTAL"].Header.Caption = CFP_GRAND_TOTAL;

            e.Layout.Bands[0].Columns["WIRE_CONFIRMATION_TOTAL"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["WIRE_CONFIRMATION_TOTAL"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["WIRE_CONFIRMATION_TOTAL"].Header.Caption = CFP_WIRE_CONF_TOTAL;
        }

#endregion

        public void SetColumnHeaderLayout()
        {
            //Initialize the header columns in the Ultragrid to a specific font, bold, height, and color.  These settings
            //from the column header will then be used when exported into a Excel worksheet.
            foreach (Infragistics.Win.UltraWinGrid.UltraGridColumn col in ultragridFinancialSummary.DisplayLayout.Bands[0].Columns)
            {
                col.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                col.Header.Appearance.BackColor = System.Drawing.Color.LightBlue; ;
                col.Header.Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
                col.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
                col.Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True;
                col.Header.Appearance.FontData.SizeInPoints = 10;
                col.Header.Appearance.FontData.Name = "Arial";
                col.Header.Appearance.ForeColor = System.Drawing.Color.Black;
                col.Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
                col.Header.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle;
            }

            //Initialize the headers and group header for the ultragrid.
            UltraGridBand band = ultragridFinancialSummary.DisplayLayout.Bands[0];
            band.RowLayoutStyle = RowLayoutStyle.GroupLayout;
            UltraGridGroup visaDPSFinTransGroup = band.Groups.Add("visaDPSFinTranGroup", CFP_VISA_DPS_FINANCIAL_TRANS);
            UltraGridGroup exceptionTransGroup = band.Groups.Add("exceptionTransGroup", CFP_EXCEPTION_TRANS);
            UltraGridGroup miscFeesGroup = band.Groups.Add("miscFeesGroup",CFP_MISCELLANEOUS_FEES);
            UltraGridGroup visaAtmReimbFeesGroup = band.Groups.Add("visaAtmReimbFeesGroup", CFP_VISA_ATM_REIMB_FEES);
            UltraGridGroup visaAtmIsaFeesGroup = band.Groups.Add("visaAtmIsaFeesGroup", CFP_VISA_ATM_ISA_FEES);
            UltraGridGroup interlinkeReimbFeesGroup = band.Groups.Add("interlinkeReimbFeesGroup", CFP_INTERLINK_REIMB_FEES);
            UltraGridGroup interlinkIsaFeesGroup = band.Groups.Add("interlinkIsaFeesGroup", CFP_INTERLINK_ISA_FEES);



            visaDPSFinTransGroup.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            visaDPSFinTransGroup.Header.Appearance.BackColor = System.Drawing.Color.LightBlue; ;
            visaDPSFinTransGroup.Header.Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
            visaDPSFinTransGroup.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            visaDPSFinTransGroup.Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True;
            visaDPSFinTransGroup.Header.Appearance.FontData.SizeInPoints = 10;
            visaDPSFinTransGroup.Header.Appearance.FontData.Name = "Arial";
            visaDPSFinTransGroup.Header.Appearance.ForeColor = System.Drawing.Color.Black;
            visaDPSFinTransGroup.Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
            visaDPSFinTransGroup.Header.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle;

            exceptionTransGroup.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            exceptionTransGroup.Header.Appearance.BackColor = System.Drawing.Color.LightBlue; ;
            exceptionTransGroup.Header.Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
            exceptionTransGroup.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            exceptionTransGroup.Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True;
            exceptionTransGroup.Header.Appearance.FontData.SizeInPoints = 10;
            exceptionTransGroup.Header.Appearance.FontData.Name = "Arial";
            exceptionTransGroup.Header.Appearance.ForeColor = System.Drawing.Color.Black;
            exceptionTransGroup.Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
            exceptionTransGroup.Header.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle;

            miscFeesGroup.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            miscFeesGroup.Header.Appearance.BackColor = System.Drawing.Color.LightBlue; ;
            miscFeesGroup.Header.Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
            miscFeesGroup.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            miscFeesGroup.Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True;
            miscFeesGroup.Header.Appearance.FontData.SizeInPoints = 10;
            miscFeesGroup.Header.Appearance.FontData.Name = "Arial";
            miscFeesGroup.Header.Appearance.ForeColor = System.Drawing.Color.Black;
            miscFeesGroup.Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
            miscFeesGroup.Header.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle;

            visaAtmReimbFeesGroup.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            visaAtmReimbFeesGroup.Header.Appearance.BackColor = System.Drawing.Color.LightBlue; ;
            visaAtmReimbFeesGroup.Header.Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
            visaAtmReimbFeesGroup.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            visaAtmReimbFeesGroup.Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True;
            visaAtmReimbFeesGroup.Header.Appearance.FontData.SizeInPoints = 10;
            visaAtmReimbFeesGroup.Header.Appearance.FontData.Name = "Arial";
            visaAtmReimbFeesGroup.Header.Appearance.ForeColor = System.Drawing.Color.Black;
            visaAtmReimbFeesGroup.Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
            visaAtmReimbFeesGroup.Header.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle;

            visaAtmIsaFeesGroup.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            visaAtmIsaFeesGroup.Header.Appearance.BackColor = System.Drawing.Color.LightBlue; ;
            visaAtmIsaFeesGroup.Header.Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
            visaAtmIsaFeesGroup.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            visaAtmIsaFeesGroup.Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True;
            visaAtmIsaFeesGroup.Header.Appearance.FontData.SizeInPoints = 10;
            visaAtmIsaFeesGroup.Header.Appearance.FontData.Name = "Arial";
            visaAtmIsaFeesGroup.Header.Appearance.ForeColor = System.Drawing.Color.Black;
            visaAtmIsaFeesGroup.Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
            visaAtmIsaFeesGroup.Header.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle;

            interlinkeReimbFeesGroup.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            interlinkeReimbFeesGroup.Header.Appearance.BackColor = System.Drawing.Color.LightBlue; ;
            interlinkeReimbFeesGroup.Header.Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
            interlinkeReimbFeesGroup.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            interlinkeReimbFeesGroup.Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True;
            interlinkeReimbFeesGroup.Header.Appearance.FontData.SizeInPoints = 10;
            interlinkeReimbFeesGroup.Header.Appearance.FontData.Name = "Arial";
            interlinkeReimbFeesGroup.Header.Appearance.ForeColor = System.Drawing.Color.Black;
            interlinkeReimbFeesGroup.Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
            interlinkeReimbFeesGroup.Header.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle;

            interlinkIsaFeesGroup.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            interlinkIsaFeesGroup.Header.Appearance.BackColor = System.Drawing.Color.LightBlue; ;
            interlinkIsaFeesGroup.Header.Appearance.BackColor2 = System.Drawing.Color.LightSkyBlue;
            interlinkIsaFeesGroup.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            interlinkIsaFeesGroup.Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True;
            interlinkIsaFeesGroup.Header.Appearance.FontData.SizeInPoints = 10;
            interlinkIsaFeesGroup.Header.Appearance.FontData.Name = "Arial";
            interlinkIsaFeesGroup.Header.Appearance.ForeColor = System.Drawing.Color.Black;
            interlinkIsaFeesGroup.Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
            interlinkIsaFeesGroup.Header.Appearance.TextVAlign = Infragistics.Win.VAlign.Middle;


            //assign the columns to the group and set the header font color for these columns to blue.
            band.Columns["VISA_DPS_FINANCIAL_TRANS"].RowLayoutColumnInfo.ParentGroup = visaDPSFinTransGroup;
            band.Columns["VISA_DPS_FINANCIAL_TRANS"].Header.Appearance.ForeColor = System.Drawing.Color.Blue;

            band.Columns["EXCEPTION_TRANS"].RowLayoutColumnInfo.ParentGroup = exceptionTransGroup;
            band.Columns["EXCEPTION_TRANS"].Header.Appearance.ForeColor = System.Drawing.Color.Blue;

            band.Columns["MISCELLANEOUS_FEES"].RowLayoutColumnInfo.ParentGroup = miscFeesGroup;
            band.Columns["MISCELLANEOUS_FEES"].Header.Appearance.ForeColor = System.Drawing.Color.Blue;

            band.Columns["VISA_ATM_REIMB_FEES"].RowLayoutColumnInfo.ParentGroup = visaAtmReimbFeesGroup;
            band.Columns["VISA_ATM_REIMB_FEES"].Header.Appearance.ForeColor = System.Drawing.Color.Blue;

            band.Columns["VISA_ATM_ISA_FEES"].RowLayoutColumnInfo.ParentGroup = visaAtmIsaFeesGroup;
            band.Columns["VISA_ATM_ISA_FEES"].Header.Appearance.ForeColor = System.Drawing.Color.Blue;

            band.Columns["INTERLINK_REIMB_FEES"].RowLayoutColumnInfo.ParentGroup = interlinkeReimbFeesGroup;
            band.Columns["INTERLINK_REIMB_FEES"].Header.Appearance.ForeColor = System.Drawing.Color.Blue;

            band.Columns["INTERLINK_ISA_FEES"].RowLayoutColumnInfo.ParentGroup = interlinkIsaFeesGroup;
            band.Columns["INTERLINK_ISA_FEES"].Header.Appearance.ForeColor = System.Drawing.Color.Blue;

            //arrange the columns in the correct order. By default when grouping, Infragistics will display the
            //grouped columns first and then the ungrouped columns last.
            //1. loop through every column and set the RowLayoutCoumnInfo.SpanX and SpanY to 1.
            for (int i = 0; i < band.Columns.Count; i++)
            {
                band.Columns[i].RowLayoutColumnInfo.SpanX = 1;
                band.Columns[i].RowLayoutColumnInfo.SpanY = 1;
            }

            band.Columns["PRC_NO"].RowLayoutColumnInfo.OriginX = 0;
            band.Columns["PRC_NO"].RowLayoutColumnInfo.SpanY = 2;
            band.Columns["REF"].RowLayoutColumnInfo.OriginX = 1;
            band.Columns["REF"].RowLayoutColumnInfo.SpanY = 2;
            band.Columns["CREDIT_UNION"].RowLayoutColumnInfo.OriginX = 2;
            band.Columns["CREDIT_UNION"].RowLayoutColumnInfo.SpanY = 2;

            visaDPSFinTransGroup.RowLayoutGroupInfo.OriginX = 3;
            visaDPSFinTransGroup.RowLayoutGroupInfo.SpanY = 1;
            exceptionTransGroup.RowLayoutGroupInfo.OriginX = 4;
            exceptionTransGroup.RowLayoutGroupInfo.SpanY = 1;
            miscFeesGroup.RowLayoutGroupInfo.OriginX = 5;
            miscFeesGroup.RowLayoutGroupInfo.SpanY = 1;
            visaAtmReimbFeesGroup.RowLayoutGroupInfo.OriginX = 6;
            visaAtmReimbFeesGroup.RowLayoutGroupInfo.SpanY = 1;
            visaAtmIsaFeesGroup.RowLayoutGroupInfo.OriginX = 7;
            visaAtmIsaFeesGroup.RowLayoutGroupInfo.SpanY = 1;
            interlinkeReimbFeesGroup.RowLayoutGroupInfo.OriginX = 8;
            interlinkeReimbFeesGroup.RowLayoutGroupInfo.SpanY = 1;
            interlinkIsaFeesGroup.RowLayoutGroupInfo.OriginX = 9;
            interlinkIsaFeesGroup.RowLayoutGroupInfo.SpanY = 1;

            band.Columns["GRAND_TOTAL"].RowLayoutColumnInfo.OriginX = 10;
            band.Columns["GRAND_TOTAL"].RowLayoutColumnInfo.SpanY = 2;
            band.Columns["WIRE_CONFIRMATION_TOTAL"].RowLayoutColumnInfo.OriginX = 11;
            band.Columns["WIRE_CONFIRMATION_TOTAL"].RowLayoutColumnInfo.SpanY = 2;



        }

        //VS4229
        /// <summary>
        /// SetDataContainer method updates the dictionary that holds the wireconfirmation. Any PRCNO that
        /// are present on this list means that their column for Wire Confirmation Total gets filled with data.
        /// </summary>
        /// <param name="inContainer"></param>
        public void SetDataContainer(DataContainer inContainer)
        {
            if (inContainer != null)
            {
                myContainer = inContainer;
                wireConfTable = myContainer.GetDictionaryFromListKeyPair();
            }

            else
            {
                //insert the default values into this table that will be used Wire Confirmation Total.
                MAPacificReportUtility.forms.SubClientIDPreferencesForm myForm = new forms.SubClientIDPreferencesForm();
                myContainer = myForm.MyDContainer;
                wireConfTable = myContainer.GetDictionaryFromListKeyPair();

                /*
                wireConfTable.Add("PRC208", true);
                wireConfTable.Add("PRC339", true);
                wireConfTable.Add("PRC347", true);
                wireConfTable.Add("PRC451", true);
                wireConfTable.Add("PRC564", true);
                wireConfTable.Add("PRC586", true);
                wireConfTable.Add("PRC600", true);
                wireConfTable.Add("PRC603", true);
                wireConfTable.Add("PRC607", true);
                wireConfTable.Add("PRC615", true);
                wireConfTable.Add("PRC665", true);
                wireConfTable.Add("PRC692", true);
                wireConfTable.Add("PRC765", true);
                wireConfTable.Add("PRC766", true);
                wireConfTable.Add("PRC823", true);
                wireConfTable.Add("PRC828", true);
                wireConfTable.Add("PRC966", true);
                wireConfTable.Add("PRC968", true);
                wireConfTable.Add("PRC211", true);
                wireConfTable.Add("PRC558", true);
                wireConfTable.Add("PRC580", true);
                wireConfTable.Add("PRC818", true);
                wireConfTable.Add("PRC819", true);
                wireConfTable.Add("PRC885", true);
                wireConfTable.Add("PRC887", true);
                wireConfTable.Add("PRC975", true);
                */
            }
        }

        public void ParseData()
        {
            //$d01 use a loop to retrieve the OV screen text. The OV screen will always have some sort of texts. Its just that the ovApi
            //has issues retrieveing that content when the host is writing to the screen and we are trying to retrieve it at the same time....
            int counter = 0;
            do
            {
                //$d01
                //More strange behaviors on 64bit OS platform.  Sometimes the ovApi.Text returns blank value when there is data on the OV screen.
                //As a work around, we will try to get the ovApi.Text untill we actually have data in there...
                apiText = ovApi.Text;
                TraceLine("ParseData datalength: " + apiText.Length + " try number: " + counter++ + "\r\n" + apiText);
                //add precaution logic so we dont get stuck in an endless loop if the ovApi.Text is not able to retrieve meeningful data or
                //something happened.
                if (counter > 10)
                    break;
            } while (apiText.Length <= 3);

            SetState(apiText);

            switch (CurrentState)
            {
                case ProcessWireConfirmationReport.WireConfReportState.ReportSelectionListState:
                    FunctionKey = EnterKey;
                    SetReportReportSelectionList();
                    break;

                case ProcessWireConfirmationReport.WireConfReportState.PageIndexSelectionListState:
                    FunctionKey = EnterKey;
                    SetPageIndexSelectionList();
                    break;

                case ProcessWireConfirmationReport.WireConfReportState.SarReportState:
                    FunctionKey = F8Key;
                    ParsePage(apiText);
                    try
                    {
                        //CurrentState may be changed...
                        if(CurrentState != WireConfReportState.FinishState)
                            ProcessFinancialSummaryData(apiText);
                    }
                    catch (Exception e2)
                    {
                        //CurrentState may be changed...
                        if(CurrentState != WireConfReportState.FinishState)
                        ProcessFinancialSummaryData(apiText);
                    }
                    break;

                case ProcessWireConfirmationReport.WireConfReportState.FinishState:
                    CurrentState = WireConfReportState.PageIndexSelectionListState;
                    FunctionKey = F3Key;
                    break;
            }

            //if (CurrentState == WireConfReportState.SarReportState)
            //{
            //    try
            //    {
            //        ProcessFinancialSummaryData(apiText);
            //    }
            //    catch (Exception e2)
            //    {
            //        ProcessFinancialSummaryData(apiText);
            //    }
            //}
        }

        /// <summary>
        /// Normally we could anticipate the state changes. However when playing a trace back; its difficult to
        /// know when the next state change is because the function key transmission does not trigger an immediate screen
        /// change. As a work around, we will have to determine the different states we are in by parsing the screen content.
        /// 
        /// </summary>
        /// <param name="inText"></param>
        public void SetState(string inText)
        {
            if (isTargetTextPresent(inText, new Regex("Report Selection List")))
            {
                CurrentState = WireConfReportState.ReportSelectionListState;
            }

            if (isTargetTextPresent(inText, new Regex("Page Index Selection List")))
            {
                CurrentState = WireConfReportState.PageIndexSelectionListState;
            }

            if (isTargetTextPresent(inText, new Regex("(SARPAGE\\s\\d*)")))
            {
                CurrentState = WireConfReportState.SarReportState;
            }
        }

        /// <summary>
        /// Select the ReportID from the Report Selection List Page. As you select the report and process them,
        /// The previous report ID will scroll off the screen.
        /// </summary>
        public void SetReportReportSelectionList()
        {
            SarPage = 0;
            //
            //check for a valid page to parse. A valid page would contain the Report Selection List keyword.
            //What we want is the "Report Selection List".
            if (isTargetTextPresent(apiText, new Regex("Report Selection List")))
            {
                //determine if we are parsing the first record on this list. If so then we can select
                //the first report on the list by placing an S into the input field for index 9.
                if (PriorPRCNO.Length <= 0)
                {
                    ovApi.InputFields[9].Text = "S";
                    PriorPRCNO = ovApi.DataFields[24].Text;
                  //  CurrentState = WireConfReportState.PageIndexSelectionListState;
                }
                //check Data Field label with that of PriorPRCNO to see if they match.
                //If they do, then we need to select the next one on the list. The previous ID has not scrolled
                //off the screen yet.
                else
                {
                    if (ovApi.DataFields[24].Text.Equals(PriorPRCNO, StringComparison.CurrentCultureIgnoreCase))
                    {
                        //Number
                        if (ovApi.InputFields.Count >= 11)
                        {
                            //Select the next row instead.
                            ovApi.InputFields[10].Text = "S";
                            //set the PriorPRCNO to point to this label that we just selected.
                            PriorPRCNO = ovApi.DataFields[32].Text;
                         //   CurrentState = WireConfReportState.PageIndexSelectionListState;
                        }
                        else
                        {
                            //we are at the end of the report selection list there is only 1 PRCNO left 
                            //and the value is the same as PriorPRCNO then we are done.
                            ConvertDataToWireReport();
                            ReportFinish = true;
                            EndofReport();
                        }
                    }
                    //set the PriorPRCNO to previous value
                    else
                    {
                        PriorPRCNO = ovApi.DataFields[24].Text;
                        //Check to see if there are anymore Report Selection List after processing the last item.
                        if(ovApi.InputFields.Count <= 10)
                        {
                            //we are at the end of the report selection list there is only 1 PRCNO left 
                            //and the value is the same as PriorPRCNO then we are done.
                            ConvertDataToWireReport();
                            ReportFinish = true;
                            EndofReport();
                        }
                    }
                }
            }
            //We are not in the correct page lets exit out and notify the user.
            else
            {
                mError = true;
                ReportFinish = true;
            }
        }

        /// <summary>
        /// Select the ALL PAGES option within the Page Index Selection List page.
        /// </summary>
        public void SetPageIndexSelectionList()
        {
            //We are trying to go back up to ReportSelectionList to work on the next Report ID...
            if (mKickOutToReportSelectionList)
            {
                mKickOutToReportSelectionList = false;
                FunctionKey = F3Key;
              //  CurrentState = WireConfReportState.ReportSelectionListState;
                return;
            }

            //
            //check for a valid page to parse. A valid page would contain the Page Index Selection List keyword.
            //What we want is the "Page Index Selection List".
            if (isTargetTextPresent(apiText, new Regex("Page Index Selection List")))
            {
                //we will be parsing all pages. Place an S into the InputFields for ALL PAGES option.
                ovApi.InputFields[2].Text = "S";

             //   CurrentState = WireConfReportState.SarReportState;
            }
            //We are not in the correct page lets exit out and notify the user.
            else
            {
                mError = true;
                ReportFinish = true;
            }
        }

        public void ParsePage(string inText)
        {
            //Set the report rundate if it hasnt been set yet.
            if (!bRunDateSet)
            {
                bRunDateSet = true;
                if (isTargetTextPresent(inText, new Regex("RUNDATE")))
                {
                    //find the index position where RUNDATE is located
                    int position = inText.IndexOf("RUNDATE ");
                    position += 7; //factor the length of string RUNDATE
                    string date = inText.Substring(position, 10).Trim();

                    //parse the date values...
                    Regex targetStr = new Regex("[0-9]{2}");
                    MatchCollection mCollection = targetStr.Matches(date);
                    if (mCollection.Count > 0)
                    {
                        for (int i = 0; i < mCollection.Count; i++)
                        {
                            RunDate += mCollection[i].Value;
                            //append the underscore
                            if( i < 2)
                                RunDate += "_";
                        }
                    }
                }
            }
            
            //
            //check for a valid page to parse. A valid page would contain the SARPAGE keyword.
            //What we want is the "SARPAGE 1" keyword and the digit accompanied with it.
            if (isTargetTextPresent(inText, new Regex("(SARPAGE\\s\\d*)")))
            {
                //check to make sure that we are processing a Financial Summary Report. If not then exit out and notify user.
                if(!isTargetTextPresent(inText, new Regex("FINANCIAL SUMMARY REPORT")))
                {
                    mError = true;
                    ReportFinish = true;
                    return;
                }
                
                //lets get the page number for SARPAGE which will help us in determining a new sub client ID is encountered or not.
                SetSarPage(inText);

                //check to see if we are done...
                if (mCurrentState == WireConfReportState.FinishState)
                {
                    //set the FunctionKey so that we can kick back to the Report Selection Screen...
                    FunctionKey = F3Key;
                    mKickOutToReportSelectionList = true;
                    return;
                }
            }
        }

        public void SetSarPage(String inText)
        {
            Regex sarPage = new Regex("(SARPAGE\\s\\d*)");
            string txt = inText;
            MatchCollection mCollection = sarPage.Matches(txt);
            if (mCollection != null && mCollection.Count > 0)
            {
                txt = mCollection[0].Value;
                txt = txt.Substring(7); //7 is the length for the word SARPAGE;
                txt.Trim();

                int page = 0;
                try
                {
                    page = System.Convert.ToInt32(txt);

                }
                catch (System.Exception ex)
                {
                    //do not update Sarpage..
                    return;
                }

                //important to check the page == SarPage first prior to updating SarPage.
                if (page == SarPage)//report is finished.
                {
                    CurrentState = WireConfReportState.FinishState;
                }

                //only update Sarpage if we have higher page numbers...
                if (page > SarPage)
                {
                    SarPage = page;
                    //new client branch...
                    subPage = 1; //reset the subpage for the new client branch...
                }

                //else
                //{
                //    subPage++;

                //    //VS4068 Financial Summary Report can have SarPage = 1 and SubPage > 1 which means that data is on SarPage 1 SubPage 1 and 
                //    //SarPage 1 and SubPage 2 is last page.
                //   // if (SarPage == 2 && subPage >= 2) //We only need the data from first 2 pages of Funancial Summary Report.
                //    if(SarPage >= 1 && subPage >= 2)
                //    {
                //        ReportFinish = true;
                //        ProcessFinancialSummaryData();
                //        EndofReport();
                //    }
                //}
            }
        }

        public void ResetData()
        {
            mError = false; 
            _sarpage = 0;
            subPage = 1;
            Data = "";
            ReportFinish = false;
            processWireItemList.Clear();
            CurrentState = WireConfReportState.ReportSelectionListState;

            mTotalVisaDPSFinancialTrans = 0;
            mTotalExceptionTrans = 0;
            mTotalMiscFees = 0;
            mTotalVisaATMREIMBFees = 0;
            mTotalVISAATMISAFees = 0;
            mTotalInterlinkREIMBFees = 0;
            mTotalInterlinkISAFees = 0;
            mTotalGrandTotal = 0;
            mTotalWireConfirmTotal = 0;
            ultragridFinancialSummary.InitializeLayout -= ultragridFinancialSummary_InitializeLayout;
            ultragridFinancialSummary.InitializeRow -= ultragridFinancialSummary_InitializeRow;
            ultragridFinancialSummary = new Infragistics.Win.UltraWinGrid.UltraGrid();
            ultragridFinancialSummary.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(ultragridFinancialSummary_InitializeRow);
            ultragridFinancialSummary.InitializeLayout += new InitializeLayoutEventHandler(ultragridFinancialSummary_InitializeLayout);
            bRunDateSet = false;
        }

        //$d02
        private void TraceLine(string message)
        {
            if (debug)
                System.Diagnostics.Trace.WriteLine(message + " -- thread ID: " + System.Threading.Thread.CurrentThread.ManagedThreadId);
        }

        /// <summary>
        /// Parse the Financial summary data in the Financial Summary report.
        /// </summary>
        /// <param name="data"></param>
        public void ProcessFinancialSummaryData(string inText)
        {
            ProcessWireConfirmationItem pwireItem = new ProcessWireConfirmationItem();
            string value = "";
            Regex pattern = new Regex("(.*\r\n)");
            MatchCollection mCollection = pattern.Matches(inText);
            MatchCollection mCollection2;
            string rowData;
            string alternateCUName = "";
            if (mCollection.Count > 0)
            {
                for (int i = 0; i < mCollection.Count; i++)
                {
                    rowData = mCollection[i].Value;
                    rowData = rowData.Trim();

                    //find the REF. value by searching for -123456789 value.
                    if(rowData.StartsWith("VISA CONFIDENTIAL:"))
                    {
                        pattern = new Regex("-[0-9]{9}");
                        mCollection2 = pattern.Matches(inText);
                        if (mCollection2.Count > 0)
                        {
                            //the value will have the negative sign in front of it. We want the value without the negative sign.
                            value = mCollection2[0].Value;
                            if(value != null)
                            {
                                value = value.Substring(1);
                                pwireItem.REF = value;
                            }
                        }

                        //calculate the alternate credit union name...
                        int startIndex = rowData.IndexOf("CLIENT USE ONLY") + 15; //need to add the length of the target string to index.
                        int endIndex = rowData.IndexOf(pwireItem.REF) - 1; //need to account for the '-' sign that was stripped out.
                        alternateCUName = rowData.Substring(startIndex, endIndex - startIndex).Trim();
                        
                    }

                    //find the PRC NO. and CREDIT UNTION Values which are on the same row as Financial Summary Report string.
                    if (rowData.Contains("FINANCIAL SUMMARY REPORT"))
                    {


                        //calculate the PRC NO value... pattern originally consisted of PRC and 3 digits
                        //$d03 -- pattern update. This needs to change from the 3 digits to be both alphanumeric letters.
                        pattern = new Regex("PRC[a-zA-Z0-9]{3}");
                        mCollection2 = pattern.Matches(rowData);
                        if (mCollection2.Count > 0)
                        {
                            value = mCollection2[0].Value;
                            if (value != null)
                            {
                                pwireItem.PRC_NO = value; 
                            }
                        }

                        //calculate the Credit Union value
                        //1st determine if the PRCNO falls into the special PRCNO table...
                        //If it is in the specialPRCNOTable, then we need to substitue the Credit Union name with a different value parsed
                        //from a different location on the report.
                        if (specialPRCNOTable.ContainsKey(pwireItem.PRC_NO))
                        {
                            //find the index location where the replacement credit union value is.
                            pwireItem.CREDIT_UNION = alternateCUName;
                        }
                        else //parse the Credit Union information
                        {
                            int index = rowData.IndexOf("-PRC");
                            if (index > 0)
                            {
                                value = rowData.Substring(0, index).Trim();
                                pwireItem.CREDIT_UNION = value;
                            }
                        }
                    }

                    if (rowData.StartsWith(CFP_VISA_DPS_FINANCIAL_TRANS))
                    {
                        //we are only interested in the Net column which has the CR/DR string appended to the value.
                        pattern = new Regex("([0-9,.]+DR)|([0-9,.]+CR)");
                        mCollection2 = pattern.Matches(rowData);
                        if (mCollection2.Count > 0)
                        {
                            value = mCollection2[0].Value;
                            if (value != null)
                            {
                                if (value.EndsWith("DR"))
                                {
                                   pwireItem.VISA_DPS_FINANCIAL_TRANS = Decimal.Parse(value.Substring(0, value.Length - 2));
                                }
                                else
                                {
                                   pwireItem.VISA_DPS_FINANCIAL_TRANS = Decimal.Parse(value.Substring(0, value.Length - 2)) * -1;
                                }

                                mTotalVisaDPSFinancialTrans += pwireItem.VISA_DPS_FINANCIAL_TRANS;
                            }
                        }
                    }

                    if (rowData.StartsWith(CFP_EXCEPTION_TRANS))
                    {
                        //we are only interested in the Net column which has the CR/DR string appended to the value.
                        pattern = new Regex("([0-9,.]+DR)|([0-9,.]+CR)");
                        mCollection2 = pattern.Matches(rowData);
                        if (mCollection2.Count > 0)
                        {
                            value = mCollection2[0].Value;
                            if (value != null)
                            {
                                if (value.EndsWith("DR"))
                                {
                                    pwireItem.EXCEPTION_TRANS = Decimal.Parse(value.Substring(0, value.Length - 2));
                                }
                                else
                                {
                                    pwireItem.EXCEPTION_TRANS = Decimal.Parse(value.Substring(0, value.Length - 2)) * -1;
                                }

                                mTotalExceptionTrans += pwireItem.EXCEPTION_TRANS;
                            }
                        }
                    }
                    if (rowData.StartsWith(CFP_MISCELLANEOUS_FEES))
                    {
                        //we are only interested in the Net column which has the CR/DR string appended to the value.
                        pattern = new Regex("([0-9,.]+DR)|([0-9,.]+CR)");
                        mCollection2 = pattern.Matches(rowData);
                        if (mCollection2.Count > 0)
                        {
                            value = mCollection2[0].Value;
                            if (value != null)
                            {
                                if (value.EndsWith("DR"))
                                {
                                    pwireItem.MISCELLANEOUS_FEES = Decimal.Parse(value.Substring(0, value.Length - 2));
                                }
                                else
                                {
                                    pwireItem.MISCELLANEOUS_FEES = Decimal.Parse(value.Substring(0, value.Length - 2)) * -1;
                                }

                                mTotalMiscFees += pwireItem.MISCELLANEOUS_FEES;
                            }
                        }
                    }

                    if (rowData.StartsWith(CFP_VISA_ATM_REIMB_FEES))
                    {
                        //we are only interested in the Net column which has the CR/DR string appended to the value.
                        pattern = new Regex("([0-9,.]+DR)|([0-9,.]+CR)");
                        mCollection2 = pattern.Matches(rowData);
                        if (mCollection2.Count > 0)
                        {
                            value = mCollection2[0].Value;
                            if (value != null)
                            {
                                if (value.EndsWith("DR"))
                                {
                                    pwireItem.VISA_ATM_REIMB_FEES = Decimal.Parse(value.Substring(0, value.Length - 2));
                                }
                                else
                                {
                                    pwireItem.VISA_ATM_REIMB_FEES = Decimal.Parse(value.Substring(0, value.Length - 2)) * -1;
                                }

                                mTotalVisaATMREIMBFees += pwireItem.VISA_ATM_REIMB_FEES;
                            }
                        }
                    }

                    if (rowData.StartsWith(CFP_VISA_ATM_ISA_FEES))
                    {
                        //we are only interested in the Net column which has the CR/DR string appended to the value.
                        pattern = new Regex("([0-9,.]+DR)|([0-9,.]+CR)");
                        mCollection2 = pattern.Matches(rowData);
                        if (mCollection2.Count > 0)
                        {
                            value = mCollection2[0].Value;
                            if (value != null)
                            {
                                if (value.EndsWith("DR"))
                                {
                                    pwireItem.VISA_ATM_ISA_FEES = Decimal.Parse(value.Substring(0, value.Length - 2));
                                }
                                else
                                {
                                    pwireItem.VISA_ATM_ISA_FEES = Decimal.Parse(value.Substring(0, value.Length - 2)) * -1;
                                }

                                mTotalVISAATMISAFees += pwireItem.VISA_ATM_ISA_FEES;
                            }
                        }
                    }

                    if (rowData.StartsWith(CFP_INTERLINK_REIMB_FEES))
                    {
                        //we are only interested in the Net column which has the CR/DR string appended to the value.
                        pattern = new Regex("([0-9,.]+DR)|([0-9,.]+CR)");
                        mCollection2 = pattern.Matches(rowData);
                        if (mCollection2.Count > 0)
                        {
                            value = mCollection2[0].Value;
                            if (value != null)
                            {
                                if (value.EndsWith("DR"))
                                {
                                    pwireItem.INTERLINK_REIMB_FEES = Decimal.Parse(value.Substring(0, value.Length - 2));
                                }
                                else
                                {
                                    pwireItem.INTERLINK_REIMB_FEES = Decimal.Parse(value.Substring(0, value.Length - 2)) * -1;
                                }

                                mTotalInterlinkREIMBFees += pwireItem.INTERLINK_REIMB_FEES;
                            }
                        }
                    }

                    if (rowData.StartsWith(CFP_INTERLINK_ISA_FEES))
                    {
                        //we are only interested in the Net column which has the CR/DR string appended to the value.
                        pattern = new Regex("([0-9,.]+DR)|([0-9,.]+CR)");
                        mCollection2 = pattern.Matches(rowData);
                        if (mCollection2.Count > 0)
                        {
                            value = mCollection2[0].Value;
                            if (value != null)
                            {
                                if (value.EndsWith("DR"))
                                {
                                    pwireItem.INTERLINK_ISA_FEES = Decimal.Parse(value.Substring(0, value.Length - 2));
                                }
                                else
                                {
                                    pwireItem.INTERLINK_ISA_FEES = Decimal.Parse(value.Substring(0, value.Length - 2)) * -1;
                                }

                                mTotalInterlinkISAFees += pwireItem.INTERLINK_ISA_FEES;
                            }
                        }
                    }
                }
                //We are finished parsing the page for the data now we need to perform some calculations
                //to determine the GrandTotal for each row.
                pwireItem.GRAND_TOTAL = pwireItem.VISA_DPS_FINANCIAL_TRANS + pwireItem.EXCEPTION_TRANS + pwireItem.MISCELLANEOUS_FEES
                                      + pwireItem.VISA_ATM_REIMB_FEES + pwireItem.VISA_ATM_ISA_FEES + pwireItem.INTERLINK_REIMB_FEES
                                      + pwireItem.INTERLINK_ISA_FEES;

                mTotalGrandTotal += pwireItem.GRAND_TOTAL;

                //the Wire Confirmation Total is not a calculated value. Instead we need to confirm if the 
                //PRC NO. value is present in the wireConfTable. If it is then we copy the value from the Grand Total to the Wire Conf Total.
                if (wireConfTable.ContainsKey(pwireItem.PRC_NO))
                {
                    pwireItem.WIRE_CONFIRMATION_TOTAL = pwireItem.GRAND_TOTAL;

                    mTotalWireConfirmTotal += pwireItem.WIRE_CONFIRMATION_TOTAL;
                }

                //we are finished with parsing this page.  Let add this pwireItem to the List processWireItem
                processWireItemList.Add(pwireItem);
            }
        }

        /// <summary>
        /// Create the list that will contain the data displayed in a Ultragrid to be exported to the CFPRecon worksheet.
        /// 
        /// </summary>
        public void ConvertDataToWireReport()
        {
            if (processWireItemList.Count > 0)
            {
                ProcessWireConfirmationItem pwireItem = new ProcessWireConfirmationItem();
                //create a new row that contains the totals...
                pwireItem.PRC_NO = "";
                pwireItem.REF = "";
                pwireItem.CREDIT_UNION = "";
                pwireItem.VISA_DPS_FINANCIAL_TRANS = mTotalVisaDPSFinancialTrans;
                pwireItem.EXCEPTION_TRANS = mTotalExceptionTrans;
                pwireItem.MISCELLANEOUS_FEES = mTotalMiscFees;
                pwireItem.VISA_ATM_REIMB_FEES = mTotalVisaATMREIMBFees;
                pwireItem.VISA_ATM_ISA_FEES = mTotalVISAATMISAFees;
                pwireItem.INTERLINK_REIMB_FEES = mTotalInterlinkREIMBFees;
                pwireItem.INTERLINK_ISA_FEES = mTotalInterlinkISAFees;
                pwireItem.GRAND_TOTAL = mTotalGrandTotal;
                pwireItem.WIRE_CONFIRMATION_TOTAL = mTotalWireConfirmTotal;

                processWireItemList.Add(pwireItem);
            }

            ultragridFinancialSummary.BindingContext = new System.Windows.Forms.BindingContext();
            ultragridFinancialSummary.DataSource = processWireItemList;
            SetColumnHeaderLayout();
        }
            
#region end of report event

        public delegate void EndofReportEventHandler();
        public event EndofReportEventHandler EndofReport;

#endregion

    }
#region ProcessWireConfirmationItem class
    public class ProcessWireConfirmationItem
    {
        private string mPrcNo = "";
        public string PRC_NO
        {
            get { return mPrcNo; }

            set
            {
                if (!mPrcNo.Equals(value))
                {
                    mPrcNo = string.IsNullOrEmpty(value) ? "" : value.Trim();
                }
            }
        }        
        
        private string mRef = "";
        public string REF
        {
            get { return mRef; }
            set
            {
                if (!mRef.Equals(value))
                {
                    mRef = string.IsNullOrEmpty(value) ? "" : value.Trim();
                }
            }
        }

        private string mCreditUnion = "";
        public string CREDIT_UNION
        {
            get { return mCreditUnion; }
            set
            {
                if (!mCreditUnion.Equals(value))
                {
                    mCreditUnion = string.IsNullOrEmpty(value) ? "" : value.Trim();
                }
            }
        }

        private decimal mVisaDPSFinTran = 0;
        public decimal VISA_DPS_FINANCIAL_TRANS
        {
            get { return mVisaDPSFinTran; }
            set
            {
                if (mVisaDPSFinTran != value)
                {
                    mVisaDPSFinTran = value;
                }
            }
        }

        private decimal mExceptionTrans = 0;
        public decimal EXCEPTION_TRANS
        {
            get { return mExceptionTrans; }
            set
            {
                if (mExceptionTrans != value)
                {
                    mExceptionTrans = value;
                }
            }
        }

        private decimal mMiscFees = 0;
        public decimal MISCELLANEOUS_FEES
        {
            get { return mMiscFees; }
            set
            {
                if (mMiscFees != value)
                {
                    mMiscFees = value;
                }
            }
        }

        private decimal mVisaAtmReimFees = 0;
        public decimal VISA_ATM_REIMB_FEES
        {
            get { return mVisaAtmReimFees; }
            set
            {
                if (mVisaAtmReimFees != value)
                {
                    mVisaAtmReimFees = value;
                }
            }
        }

        private decimal mVisaAtmIsaFees = 0;
        public decimal VISA_ATM_ISA_FEES
        {
            get { return mVisaAtmIsaFees; }
            set
            {
                if (mVisaAtmIsaFees != value)
                {
                    mVisaAtmIsaFees = value;
                }
            }
        }

        private decimal mInterlinkReimbFees = 0;
        public decimal INTERLINK_REIMB_FEES
        {
            get { return mInterlinkReimbFees; }
            set
            {
                if (mInterlinkReimbFees != value)
                {
                    mInterlinkReimbFees = value;
                }
            }
        }

        private decimal mInterlinkIsaFees = 0;
        public decimal INTERLINK_ISA_FEES
        {
            get { return mInterlinkIsaFees; }
            set
            {
                if (mInterlinkIsaFees != value)
                {
                    mInterlinkIsaFees = value;
                }
            }
        }

        private decimal mGrandTotal = 0;
        public decimal GRAND_TOTAL
        {
            get { return mGrandTotal; }
            set
            {
                if (mGrandTotal != value)
                {
                    mGrandTotal = value;
                }
            }
        }

        private decimal mWireConfirmationTotal;
        public decimal WIRE_CONFIRMATION_TOTAL
        {
            get { return mWireConfirmationTotal; }
            set
            {
                if (mWireConfirmationTotal != value)
                {
                    mWireConfirmationTotal = value;
                }
            }
        }

    }
#endregion
}

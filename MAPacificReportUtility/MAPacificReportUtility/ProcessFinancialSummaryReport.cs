using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Infragistics.Win.UltraWinGrid;
using Infragistics.Win;
using OvDotNet;

namespace MAPacificReportUtility
{
    class ProcessFinancialSummaryReport : ProcessReport
    {      
        public ProcessFinancialSummaryReport(OvDotNet.OvDotNetApi inApi) //Infragistics.Win.UltraWinGrid.UltraGrid ultragrid)
        {
            ovApi = inApi;
            ultragridFinancialSummary = new Infragistics.Win.UltraWinGrid.UltraGrid();

            FontData fd1 = ultragridFinancialSummary.DisplayLayout.Appearance.FontData;
            fd1.Name = "Arial";
            ultragridFinancialSummary.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(ultragridFinancialSummary_InitializeRow);
            ultragridFinancialSummary.InitializeLayout += new InitializeLayoutEventHandler(ultragridFinancialSummary_InitializeLayout);
            financialSummaryItemList = new List<FinancialSummaryItem>();
            cfpReconItemList = new List<CFPReconItem>();
        }


#region properties

        private OvDotNet.OvDotNetApi ovApi;
        private string apiText = "";
        //flag to turn on/off Trace statements
        bool debug = false;

        private decimal mWireReportFeesAndAdjustments = 0;
        private decimal mTotalCardActivity = 0;
        private decimal mTotalLoadUnloadActivity = 0;
        private decimal mTotalChangeFeeToIncome = 0;
        private decimal mTotalWireForToday = 0;
        private decimal mTotalChangeToCFP = 0;
        private decimal mEndOfDayBalance = 0;


        private const string CFP_BEGINNING_BALANCE = "CFP BEGINNING BALANCE";
        private const string CFP_EXCEPTION_TRANSACTION = "Exception Transaction";
        private const string CFP_MISCELLANEOUS_FEES = "MISCELLANEOUS FEES";
        private const string CFP_VISA_ATM_REIMB_FEES = "VISA/ATM REIMB FEES";
        private const string CFP_VISA_ATM_ISA_FEES = "VISA/ATM ISA FEES";
        private const string CFP_INTERLINK_REIMB_FEES = "INTERLINK REIMB FEES";
        private const string CFP_INTERLINK_ISA_FEES = "INTERLINK ISA FEES";
        private const string CFP_WIRE_REPORT_FEES_AND_ADJUSTMENTS = "WIRE REPORT FEES AND ADJUSTMENTS";
        private const string CFP_LOAD_DISPUTES = "LOAD DISPUTES";
        private const string CFP_PURCHASES_QUASI_CASH = "PURCHASES/QUASI CASH";
        private const string CFP_PURCHASES_WITH_CASH_BACK = "PURCHASES WITH CASH BACK";
        private const string CFP_ACCOUNT_FUNDING = "ACCOUNT FUNDING";
        private const string CFP_AFT_AA_PP = "AFT - AA/PP";
        private const string CFP_PURCHASE_RETURNS = "PURCHASE RETURNS";
        private const string CFP_MANUAL_CASH = "MANUAL CASH";
        private const string CFP_ATM_CASH = "ATM CASH";
        private const string CFP_EMERGENCY_CASH = "EMERGENCY CASH";
        private const string CFP_BALANCE_INQUIRIES = "BALANCE INQUIRIES";
        private const string CFP_TOTAL_CARD_ACTIVITY = "TOTAL CARD ACTIVITY";
        private const string CFP_LOADS_FI_FUNDS_TRANSFER = "LOADS/FI FUNDS TRANSFER";
        private const string CFP_LOADS_MERCHANT_POS_FUNDING_READYLINK = "LOADS/MERCH POS FUNDING";
        private const string CFP_UNLOADS_FI_FUNDS_TRANSFER = "UNLOADS/FI FUNDS TRNSFER";
        private const string CFP_RELOADS_FI_FUNDS_TRANSFER = "RELOADS/FI FUNDS TRNSFER";
        //VS6160 Add ACH DIRECT DEPOSIT and RELOADS/MONEY TSFR RCVD
        private const string CFP_ACH_DIRECT_DEPOSIT = "ACH DIRECT DEPOSIT";
        private const string CFP_RELOADS_MONEY_TRSFR_RCVD = "RELOADS/MONEY TSFR RCVD";
        private const string CFP_RELOADS_BYPASS = "RELOADS/BY-PASS";
        private const string CFP_UNLOADS_BYPASS = "UNLOADS/BY-PASS";
        private const string CFP_RELOADS_MERCHANT_ACQUIRER = "RELOADS/MERCHANT ACQ";
        private const string CFP_TOTAL_LOAD_UNLOAD_ACTIVITY = "TOTAL LOAD/UNLOAD ACTIVITY";
        private const string CFP_TRANSACTION_DISPUTES = "TRANSACTION DISPUTES";
        private const string CFP_MERCHANT_ADJUSTMENTS = "MERCHANT ADJUSTMENTS";
        private const string CFP_MANUAL_ADJUSTMENTS = "MANUAL ADJUSTMENTS";
        private const string CFP_CLOSED_FOR_ESCHEATMENT = "CLOSED FOR ESCHEATMENT";
        private const string CFP_CARD_PGM_FEES = "CARD PGM FEES";
        private const string CFP_CARD_PGM_FEE_ADJUSTMENTS = "CARD PGM FEE ADJUSTMENTS";
        private const string CFP_TOTAL_CHANGE_TO_FEE_INCOME_ACCOUNT = "TOTAL CHANGE TO FEE INCOME ACCOUNT";
        private const string CFP_TOTAL_WIRE_FOR_TODAY = "TOTAL WIRE FOR TODAY";
        private const string CFP_TOTAL_CHANGE_TO_CFP = "TOTAL CHANGE TO CFP";
        private const string CFP_END_OF_DAY_BALANCE = "END OF DAY BALANCE (s/b $500 less than balance on TranZact due to original $5000 prefund)";
        //VS4637 Add Loads Bypass, Loads Merch Acq, and Unloads Merch Acq
        private const string CFP_LOADS_BYPASS = "LOADS/BY-PASS";
        private const string CFP_LOADS_MERCHANT_ACQUIRER = "LOADS/MERCHANT ACQ";
        private const string CFP_UNLOADS_MERCHANT_ACQUIRER = "UNLOADS/MERCHANT ACQ";
        //VS4649 Add CLOSED FOR NEG BAL
        private const string CFP_CLOSED_FOR_NEG_BAL = "CLOSED FOR NEGATIVE BAL";
        /// <summary>
        /// Global Summary Report data
        /// </summary>
        private GlobalSummaryBranch mSummaryBranch;
        public GlobalSummaryBranch GSummaryBranch
        {
            set { mSummaryBranch = value; }
            get { return mSummaryBranch; }
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

        private List<FinancialSummaryItem> financialSummaryItemList;

        private List<CFPReconItem> cfpReconItemList;
        public List<CFPReconItem> CFPReconItemList
        {
            get { return cfpReconItemList; }
        }

        private Infragistics.Win.UltraWinGrid.UltraGrid ultragridFinancialSummary;
        public Infragistics.Win.UltraWinGrid.UltraGrid UltraGridFinancialSummary
        {
            get { return ultragridFinancialSummary; }
        }

#endregion

#region ultragrid events

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

            e.Layout.Bands[0].Columns["Description"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["Debit1"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["Debit1"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["Credit1"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["Credit1"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["Debit2"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["Debit2"].CellAppearance.TextVAlign = VAlign.Middle;
            e.Layout.Bands[0].Columns["Credit2"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["Credit2"].CellAppearance.TextVAlign = VAlign.Middle;


        }

        /// <summary>
        /// InitializeRow event allows us to set various cell appearance for cells in that specific row.  In our case we are
        /// setting Font blod, Font type Arial, and cell background color to Yellow.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultragridFinancialSummary_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
        {
            //Add bold attributes to certain row data so that they are easy to read...
            if (e.Row.Cells["Description"].Value.ToString().Equals(CFP_BEGINNING_BALANCE, StringComparison.CurrentCultureIgnoreCase))
            {
                e.Row.Cells["Description"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["Description"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["Description"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["Description"].Appearance.ForeColor = System.Drawing.Color.Black;

                if (((decimal)(e.Row.Cells["Debit1"].Value) != 0 ))
                {
                    e.Row.Cells["Debit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit1"].Value) != 0))
                {
                    e.Row.Cells["Credit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Debit2"].Value) != 0))
                {
                    e.Row.Cells["Debit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit2"].Value) != 0))
                {
                    e.Row.Cells["Credit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }
            }

            if (e.Row.Cells["Description"].Value.ToString().Equals(CFP_WIRE_REPORT_FEES_AND_ADJUSTMENTS, StringComparison.CurrentCultureIgnoreCase))
            {
                e.Row.Cells["Description"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["Description"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["Description"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["Description"].Appearance.ForeColor = System.Drawing.Color.Black;

                if (((decimal)(e.Row.Cells["Debit1"].Value) != 0))
                {
                    e.Row.Cells["Debit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit1"].Value) != 0))
                {
                    e.Row.Cells["Credit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Debit2"].Value) != 0))
                {
                    e.Row.Cells["Debit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit2"].Value) != 0))
                {
                    e.Row.Cells["Credit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }
            }

            if (e.Row.Cells["Description"].Value.ToString().Equals(CFP_TOTAL_CARD_ACTIVITY, StringComparison.CurrentCultureIgnoreCase))
            {
                e.Row.Cells["Description"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["Description"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["Description"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["Description"].Appearance.ForeColor = System.Drawing.Color.Black;

                if (((decimal)(e.Row.Cells["Debit1"].Value) != 0))
                {
                    e.Row.Cells["Debit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit1"].Value) != 0))
                {
                    e.Row.Cells["Credit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Debit2"].Value) != 0))
                {
                    e.Row.Cells["Debit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit2"].Value) != 0))
                {
                    e.Row.Cells["Credit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }
            }

            if (e.Row.Cells["Description"].Value.ToString().Equals(CFP_TOTAL_LOAD_UNLOAD_ACTIVITY, StringComparison.CurrentCultureIgnoreCase))
            {
                e.Row.Cells["Description"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["Description"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["Description"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["Description"].Appearance.ForeColor = System.Drawing.Color.Black;

                if (((decimal)(e.Row.Cells["Debit1"].Value) != 0))
                {
                    e.Row.Cells["Debit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit1"].Value) != 0))
                {
                    e.Row.Cells["Credit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Debit2"].Value) != 0))
                {
                    e.Row.Cells["Debit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit2"].Value) != 0))
                {
                    e.Row.Cells["Credit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }
            }

            if (e.Row.Cells["Description"].Value.ToString().Equals(CFP_TOTAL_CHANGE_TO_FEE_INCOME_ACCOUNT, StringComparison.CurrentCultureIgnoreCase))
            {
                e.Row.Cells["Description"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["Description"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["Description"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["Description"].Appearance.ForeColor = System.Drawing.Color.Black;

                if (((decimal)(e.Row.Cells["Debit1"].Value) != 0))
                {
                    e.Row.Cells["Debit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit1"].Value) != 0))
                {
                    e.Row.Cells["Credit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Debit2"].Value) != 0))
                {
                    e.Row.Cells["Debit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit2"].Value) != 0))
                {
                    e.Row.Cells["Credit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }
            }

            if (e.Row.Cells["Description"].Value.ToString().Equals(CFP_TOTAL_WIRE_FOR_TODAY, StringComparison.CurrentCultureIgnoreCase))
            {
                e.Row.Cells["Description"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["Description"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["Description"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["Description"].Appearance.ForeColor = System.Drawing.Color.Black;

                if (((decimal)(e.Row.Cells["Debit1"].Value) != 0))
                {
                    e.Row.Cells["Debit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit1"].Value) != 0))
                {
                    e.Row.Cells["Credit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Debit2"].Value) != 0))
                {
                    e.Row.Cells["Debit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit2"].Value) != 0))
                {
                    e.Row.Cells["Credit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

            }

            if (e.Row.Cells["Description"].Value.ToString().Equals(CFP_TOTAL_CHANGE_TO_CFP, StringComparison.CurrentCultureIgnoreCase))
            {
                e.Row.Cells["Description"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["Description"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["Description"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["Description"].Appearance.ForeColor = System.Drawing.Color.Black;

                if (((decimal)(e.Row.Cells["Debit1"].Value) != 0))
                {
                    e.Row.Cells["Debit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit1"].Value) != 0))
                {
                    e.Row.Cells["Credit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Debit2"].Value) != 0))
                {
                    e.Row.Cells["Debit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit2"].Value) != 0))
                {
                    e.Row.Cells["Credit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

            }

            if (e.Row.Cells["Description"].Value.ToString().Equals(CFP_END_OF_DAY_BALANCE, StringComparison.CurrentCultureIgnoreCase))
            {
                e.Row.Cells["Description"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["Description"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["Description"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["Description"].Appearance.ForeColor = System.Drawing.Color.Black;

                if (((decimal)(e.Row.Cells["Debit1"].Value) != 0))
                {
                    e.Row.Cells["Debit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit1"].Value) != 0))
                {
                    e.Row.Cells["Credit1"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit1"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Debit2"].Value) != 0))
                {
                    e.Row.Cells["Debit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Debit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }

                if (((decimal)(e.Row.Cells["Credit2"].Value) != 0))
                {
                    e.Row.Cells["Credit2"].Appearance.FontData.Bold = DefaultableBoolean.True;
                    e.Row.Cells["Credit2"].Appearance.BackColor = System.Drawing.Color.Yellow;
                }
            } 
        }
#endregion

        public void SetColumnHeaderLayout()
        {
            //Initialize the header columns in the Ultragrid to a specific font, bold, height, and color.  These settings
            //from the column header will then be used when exported into a Excel worksheet.
            foreach (Infragistics.Win.UltraWinGrid.UltraGridColumn col in ultragridFinancialSummary.DisplayLayout.Bands[0].Columns)
            {
                col.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                col.Header.Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                col.Header.Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                col.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
                col.Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True;
                col.Header.Appearance.FontData.SizeInPoints = 10;
                col.Header.Appearance.FontData.Name = "Arial";
                col.Header.Appearance.ForeColor = System.Drawing.Color.White;
                col.Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
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

            //
            //check for a valid page to parse. A valid page would contain the SARPAGE keyword.
            //What we want is the "SARPAGE 1" keyword and the digit accompanied with it.
            if (isTargetTextPresent(apiText, new Regex("(SARPAGE\\s\\d*)")))
            {
                //check to make sure that we are processing a Financial Summary Report. If not then exit out and notify user.
                if(!isTargetTextPresent(apiText, new Regex("FINANCIAL SUMMARY REPORT")))
                {
                    mError = true;
                    ReportFinish = true;
                    return;
                }
                
                //lets get the page number for SARPAGE which will help us in determining a new sub client ID is encountered or not.
                SetSarPage(apiText);

                //Get the transaction info...
                try
                {
                    //SetTransactionType(ovApi.CrtGet(6, 0, (ovApi.Rows - 6) * ovApi.Columns));
                    //$d01 Instead of calling the CrtGet which can be flaky in 64bit OS... we just get the Text that was retrieved earlier.
                    SetTransactionType(apiText);
                }
                catch (Exception e2)
                {
                    string text = ovApi.CrtGet(6, 0, (ovApi.Rows - 6) * ovApi.Columns);
                    SetTransactionType(text);
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
                //only update Sarpage if we have higher page numbers...
                if (page > SarPage)
                {
                    SarPage = page;
                    //new client branch...
                    subPage = 1; //reset the subpage for the new client branch...
                }
         //       else if (page < SarPage)//somehow we went to the previous screen...
         //       {
         //           isPreviousSarPage = true;
         //       }
                else
                {
                    subPage++;

                    //VS4068 Financial Summary Report can have SarPage = 1 and SubPage > 1 which means that data is on SarPage 1 SubPage 1 and 
                    //SarPage 1 and SubPage 2 is last page.
                   // if (SarPage == 2 && subPage >= 2) //We only need the data from first 2 pages of Funancial Summary Report.
                    if(SarPage >= 1 && subPage >= 2)
                    {
                        ReportFinish = true;
                        ProcessFinancialSummaryData();
                        EndofReport();
                    }
                }
            }
        }

        public void ResetData()
        {
            mError = false; 
            _sarpage = 0;
            subPage = 1;
            Data = "";
            ReportFinish = false;
            CFPReconItemList.Clear();
            financialSummaryItemList.Clear();
            mWireReportFeesAndAdjustments = 0;
            mTotalCardActivity = 0;
            mTotalLoadUnloadActivity = 0;
            mTotalChangeFeeToIncome = 0;
            mTotalWireForToday = 0;
            mTotalChangeToCFP = 0;
            mEndOfDayBalance = 0;
            ultragridFinancialSummary.InitializeLayout -= ultragridFinancialSummary_InitializeLayout;
            ultragridFinancialSummary.InitializeRow -= ultragridFinancialSummary_InitializeRow;
            ultragridFinancialSummary = new Infragistics.Win.UltraWinGrid.UltraGrid();
            ultragridFinancialSummary.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(ultragridFinancialSummary_InitializeRow);
            ultragridFinancialSummary.InitializeLayout += new InitializeLayoutEventHandler(ultragridFinancialSummary_InitializeLayout);
        }

        //$d02
        private void TraceLine(string message)
        {
            if (debug)
                System.Diagnostics.Trace.WriteLine(message + " -- thread ID: " + System.Threading.Thread.CurrentThread.ManagedThreadId);
        }

        public void SetTransactionType(String inText)
        {
            //VS4068 Financial Summary Report can sometimes have SarPage = 1 and subPage = 2... in addition to subPage = 1 and SarPage = 2.
            //find instance of REPORT ID so we can figure out where to start retrieveing the data...
            int index = inText.LastIndexOf("DESCRIPTION");
            Data += inText.Substring(index);
            Data += Environment.NewLine;
        }

        /// <summary>
        /// Parse the Financial summary data in the Financial Summary report.
        /// </summary>
        /// <param name="data"></param>
        public void ProcessFinancialSummaryData()
        {
            Regex pattern = new Regex("(.*\r\n)");
            MatchCollection mCollection = pattern.Matches(Data);
            MatchCollection mCollection2;
            string rowData;
            FinancialSummaryItem fSumItem;
            string tmp;
            if (mCollection.Count > 0)
            {
                for (int i = 0; i < mCollection.Count; i++)
                {
                    rowData = mCollection[i].Value;
                    rowData = rowData.Trim();

                    fSumItem = new FinancialSummaryItem();

                    if (rowData.StartsWith("DESCRIPTION")) //We do not want the column header info from the data...
                    {
                        i++; //advance ahead another row because we don't need to deal with the column headers..
                        continue;
                    }

                    if (rowData.StartsWith("TOTALS"))
                    {
                        //usually after TOTALS row is a row with just underscores and no data so we can skip it.
                        i++;
                    }
                    else //we have good data to work with...
                    {
                        //ParseDescription(rowData);

                        //parse out the row information and store them into the cfpItem object
                        //1. get the description first...
                        pattern = new Regex("[\\D]+");
                        mCollection2 = pattern.Matches(rowData);
                        if (mCollection2.Count > 0)
                        {
                            fSumItem.Description = mCollection2[0].Value;
                            fSumItem.Description.Trim();
                        }

                        //After getting the Description we now need to get the other information...

                        pattern = new Regex("\\s[A-Z0-9,.-]+");
                        mCollection2 = pattern.Matches(rowData.Substring(fSumItem.Description.Length));
                        if (mCollection2.Count > 0)
                        {
                            //2. get the debit data
                            fSumItem.Debit = Decimal.Parse(mCollection2[0].Value);
                            //3. get the credit data
                            fSumItem.Credit = Decimal.Parse(mCollection2[1].Value);
                            //4. get the Net data
                            tmp = mCollection2[2].Value;
                            fSumItem.Net = Decimal.Parse(tmp.Substring(0, (tmp.Length - 2)));
                            fSumItem.DBCR = tmp.Substring(tmp.Length - 2);
                            //5. get the Report ID data
                            fSumItem.ReportID = mCollection2[3].Value;

                            //add the fSumItem to the list...
                            financialSummaryItemList.Add(fSumItem);
                        }

                    }
                }
            }
        }

        /// <summary>
        /// Create the list that will contain the data displayed in a Ultragrid to be exported to the CFPRecon worksheet.
        /// 
        /// </summary>
        public void ConvertDataToCFPRecon()
        {
            if (financialSummaryItemList.Count > 0 && GSummaryBranch.TransactionList.Count > 0)
            {
                
#region CFP Beginning Balance                
                CFPReconItem cfpItem;
                FinancialSummaryItem fItem;
                //Add first entry for CFP BEGINNING BALANCE
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_BEGINNING_BALANCE;
                cfpItem.Credit1 = GSummaryBranch.FundsPoolStatusList[0].FundsPoolBalance;
                CFPReconItemList.Add(cfpItem);

                //calculate EndOfDayBalance...
                mEndOfDayBalance += cfpItem.Credit1;

                //Add the entries from Financial Summary Report
                cfpItem = new CFPReconItem();
                fItem = GetFinancialSummaryItem(CFP_EXCEPTION_TRANSACTION);
                if(fItem != null)
                {
                    cfpItem.Description = fItem.Description + " - " + fItem.ReportID;
                    if (fItem.DBCR == "DR")
                    {
                        cfpItem.Credit2 = fItem.Net;
                        mTotalWireForToday += cfpItem.Credit2;
                    }
                    else
                    {
                        cfpItem.Debit2 = (fItem.Net * -1);
                        mTotalWireForToday += cfpItem.Debit2;
                    }

                    CFPReconItemList.Add(cfpItem);
                }

                cfpItem = new CFPReconItem();
                fItem = GetFinancialSummaryItem(CFP_MISCELLANEOUS_FEES);
                if (fItem != null)
                {
                    cfpItem.Description = fItem.Description + " - " + fItem.ReportID;
                    if (fItem.DBCR == "DR")
                    {
                        cfpItem.Credit2 = fItem.Net;
                        mWireReportFeesAndAdjustments += cfpItem.Credit2;
                    }
                    else
                    {
                        cfpItem.Debit2 = (fItem.Net * -1);
                        mWireReportFeesAndAdjustments += cfpItem.Debit2;
                    }

                    CFPReconItemList.Add(cfpItem);
                }

                cfpItem = new CFPReconItem();
                fItem = GetFinancialSummaryItem(CFP_VISA_ATM_REIMB_FEES);
                if (fItem != null)
                {
                    cfpItem.Description = fItem.Description + " - " + fItem.ReportID;
                    if (fItem.DBCR == "DR")
                    {
                        cfpItem.Credit2 = fItem.Net;
                        mWireReportFeesAndAdjustments += cfpItem.Credit2;
                    }
                    else
                    {
                        cfpItem.Debit2 = (fItem.Net * -1);
                        mWireReportFeesAndAdjustments += cfpItem.Debit2;
                    }

                    CFPReconItemList.Add(cfpItem);
                }

                cfpItem = new CFPReconItem();
                fItem = GetFinancialSummaryItem(CFP_VISA_ATM_ISA_FEES);
                if (fItem != null)
                {
                    cfpItem.Description = fItem.Description + " - " + fItem.ReportID;
                    if (fItem.DBCR == "DR")
                    {
                        cfpItem.Credit2 = fItem.Net;
                        mWireReportFeesAndAdjustments += cfpItem.Credit2;
                    }
                    else
                    {
                        cfpItem.Debit2 = (fItem.Net * -1);
                        mWireReportFeesAndAdjustments += cfpItem.Debit2;
                    }

                    CFPReconItemList.Add(cfpItem);
                }


                cfpItem = new CFPReconItem();
                fItem = GetFinancialSummaryItem(CFP_INTERLINK_REIMB_FEES);
                if (fItem != null)
                {
                    cfpItem.Description = fItem.Description + " - " + fItem.ReportID;
                    if (fItem.DBCR == "DR")
                    {
                        cfpItem.Credit2 = fItem.Net;
                        mWireReportFeesAndAdjustments += cfpItem.Credit2;
                    }
                    else
                    {
                        cfpItem.Debit2 = (fItem.Net * -1);
                        mWireReportFeesAndAdjustments += cfpItem.Debit2;
                    }

                    CFPReconItemList.Add(cfpItem);
                }

                //Note that the Interlink ISA Fees may or may not be present in the report.
                cfpItem = new CFPReconItem();
                fItem = GetFinancialSummaryItem(CFP_INTERLINK_ISA_FEES);
                if (fItem != null)
                {
                    cfpItem.Description = fItem.Description + " - " + fItem.ReportID;
                    if (fItem.DBCR == "DR")
                    {
                        cfpItem.Credit2 = fItem.Net;
                        mWireReportFeesAndAdjustments += cfpItem.Credit2;
                    }
                    else
                    {
                        cfpItem.Debit2 = (fItem.Net * -1);
                        mWireReportFeesAndAdjustments += cfpItem.Debit2;
                    }

                    CFPReconItemList.Add(cfpItem);
                }
#endregion

#region Wire Report Fees And Adjustment
                //create row data for Wire Report Fees and Adjustments.
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_WIRE_REPORT_FEES_AND_ADJUSTMENTS;
                if (mWireReportFeesAndAdjustments > 0)
                {
                    cfpItem.Debit2 = mWireReportFeesAndAdjustments;
                }
                else
                {
                    cfpItem.Credit2 = mWireReportFeesAndAdjustments;
                }
                cfpReconItemList.Add(cfpItem);

                //add wirereportfeesandadjustments to totalchangefeetoincome...
                mTotalChangeFeeToIncome += mWireReportFeesAndAdjustments;
                mTotalWireForToday += mWireReportFeesAndAdjustments;

                //write the data from Visa DPS Global Summary under Wire Report Fees and Adjustments row...
                //Get Load Disputes
                TransactionType aTransactionItem;
                aTransactionItem = GetTransactionItem(CFP_LOAD_DISPUTES);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_LOAD_DISPUTES;
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalCardActivity += cfpItem.Debit1;
                    }
                    else if(aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalCardActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get Purchase/Quasi Cash
                aTransactionItem = GetTransactionItem(CFP_PURCHASES_QUASI_CASH);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_PURCHASES_QUASI_CASH;
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalCardActivity += cfpItem.Debit1;
                    }
                    else if(aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalCardActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get Purchases with cash back
                aTransactionItem = GetTransactionItem(CFP_PURCHASES_WITH_CASH_BACK);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_PURCHASES_WITH_CASH_BACK;
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalCardActivity += cfpItem.Debit1;
                    }
                    else if(aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalCardActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get account funding
                aTransactionItem = GetTransactionItem(CFP_ACCOUNT_FUNDING);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_ACCOUNT_FUNDING;
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalCardActivity += cfpItem.Debit1;
                    }
                    else if(aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalCardActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get AFT - AA/PP
                aTransactionItem = GetTransactionItem(CFP_AFT_AA_PP);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_AFT_AA_PP;
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalCardActivity += cfpItem.Debit1;
                    }
                    else if(aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalCardActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get Purchase returns
                aTransactionItem = GetTransactionItem(CFP_PURCHASE_RETURNS);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_PURCHASE_RETURNS;
                if (aTransactionItem != null)
                {


                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalCardActivity += cfpItem.Debit1;
                    }
                    else if(aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalCardActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get Manual Cash
                aTransactionItem = GetTransactionItem(CFP_MANUAL_CASH);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_MANUAL_CASH;
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalCardActivity += cfpItem.Debit1;
                    }
                    else if(aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalCardActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

               //Get atm cash
                aTransactionItem = GetTransactionItem(CFP_ATM_CASH);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_ATM_CASH;
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalCardActivity += cfpItem.Debit1;
                    }
                    else if(aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalCardActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

              //Get Emergency Cash
                aTransactionItem = GetTransactionItem(CFP_EMERGENCY_CASH);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_EMERGENCY_CASH;
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalCardActivity += cfpItem.Debit1;
                    }
                    else if(aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalCardActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

              //Get Balance Inquiries
                aTransactionItem = GetTransactionItem(CFP_BALANCE_INQUIRIES);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_BALANCE_INQUIRIES;
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalCardActivity += cfpItem.Debit1;
                    }
                    else if(aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalCardActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);
#endregion

#region Total Card Activity
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_TOTAL_CARD_ACTIVITY;
                //value is always positive...
                cfpItem.Credit2 = Math.Abs(mTotalCardActivity);
                cfpReconItemList.Add(cfpItem);

                mTotalWireForToday += cfpItem.Credit2;
                //Changed that calculation for Total Change To CFP to subtract mTotalCardActivity value.
                mTotalChangeToCFP -= cfpItem.Credit2;

                //Get Loads Fi Funds Transfer
                aTransactionItem = GetTransactionItem(CFP_LOADS_FI_FUNDS_TRANSFER);
                cfpItem = new CFPReconItem();
                cfpItem.Description = "LOADS FI FUNDS TRANSFER";
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalLoadUnloadActivity += cfpItem.Debit1;
                    }
                    else if (aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalLoadUnloadActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get Loads Merchant Pos Funding...
                aTransactionItem = GetTransactionItem(CFP_LOADS_MERCHANT_POS_FUNDING_READYLINK);
                cfpItem = new CFPReconItem();
                cfpItem.Description = "LOADS MERCHANT POS FUNDING";
                if (aTransactionItem != null)
                {
                    //value shall always be a credit amount placed into the first set debit/credit column. The reverse value
                    //shall be placed onto the debit side of the second set of debit credit column.
                    cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                    cfpItem.Debit2 = (aTransactionItem.TransactionAmount * -1);
                    mTotalLoadUnloadActivity += cfpItem.Credit1;
                    mTotalWireForToday += cfpItem.Debit2;
                }
                CFPReconItemList.Add(cfpItem);

                //Get Unloads FI Funds Transfer
                aTransactionItem = GetTransactionItem(CFP_UNLOADS_FI_FUNDS_TRANSFER);
                cfpItem = new CFPReconItem();
                cfpItem.Description = "UNLOADS FI FUNDS TRANSFER";
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalLoadUnloadActivity += cfpItem.Debit1;
                    }
                    else if (aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalLoadUnloadActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get Reloads FI Funds Transfer
                aTransactionItem = GetTransactionItem(CFP_RELOADS_FI_FUNDS_TRANSFER);
                cfpItem = new CFPReconItem();
                cfpItem.Description = "RELOADS FI FUNDS TRANSFER";
                if (aTransactionItem != null)
                {
                    //If transaction is a Debit in the Global Summary Branch then it is a Debit in the CFPRecon report also. However
                    //the amount will be a negative number instead!.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalLoadUnloadActivity += cfpItem.Debit1;
                    }
                    else if (aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalLoadUnloadActivity += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //VS6159 Add ACH_DIRECT_DEPOSIT
                //Amount should be posted to “Credit1” (positive) column and “Debit2” (negative) column
                aTransactionItem = GetTransactionItem(CFP_ACH_DIRECT_DEPOSIT);
                cfpItem = new CFPReconItem();
                cfpItem.Description = TransactionType.TransactionOption.ACH_DIRECT_DEPOSIT;
                if (aTransactionItem != null)
                {
                    //value shall always be a credit amount placed into the first set debit/credit column. The reverse value
                    //shall be placed onto the debit side of the second set of debit credit column.
                    cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                    cfpItem.Debit2 = (aTransactionItem.TransactionAmount * -1);
                    mTotalLoadUnloadActivity += cfpItem.Credit1;
                    mTotalWireForToday += cfpItem.Debit2;
                }
                CFPReconItemList.Add(cfpItem);

                //VS6159 ADD RELOADS/MONEY TSFR RCVD
                //Amount should be posted to “Credit1” (positive) column and “Debit2” (negative) column
                aTransactionItem = GetTransactionItem(CFP_RELOADS_MONEY_TRSFR_RCVD);
                cfpItem = new CFPReconItem();
                cfpItem.Description = TransactionType.TransactionOption.RELOADS_MONEY_TRSFR_RCVD;
                if (aTransactionItem != null)
                {
                    //value shall always be a credit amount placed into the first set debit/credit column. The reverse value
                    //shall be placed onto the debit side of the second set of debit credit column.
                    cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                    cfpItem.Debit2 = (aTransactionItem.TransactionAmount * -1);
                    mTotalLoadUnloadActivity += cfpItem.Credit1;
                    mTotalWireForToday += cfpItem.Debit2;
                }
                CFPReconItemList.Add(cfpItem);

                //VS4637 add Loads Bypass
                aTransactionItem = GetTransactionItem(CFP_LOADS_BYPASS);
                cfpItem = new CFPReconItem();
                cfpItem.Description = "LOADS BYPASS";
                if (aTransactionItem != null)
                {
                    cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                    cfpItem.Debit2 = (aTransactionItem.TransactionAmount * -1);
                    mTotalLoadUnloadActivity += cfpItem.Credit1;
                    mTotalWireForToday += cfpItem.Debit2;
                }
                CFPReconItemList.Add(cfpItem);

                //Get Reloads Bypass and Loads Merch Pos Funding
                aTransactionItem = GetTransactionItem(CFP_RELOADS_BYPASS);
                cfpItem = new CFPReconItem();
                cfpItem.Description = "RELOADS BYPASS";
                if (aTransactionItem != null)
                {
                    cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                    cfpItem.Debit2 = (aTransactionItem.TransactionAmount * -1);
                    mTotalLoadUnloadActivity += cfpItem.Credit1;
                    mTotalWireForToday += cfpItem.Debit2;
                }
                CFPReconItemList.Add(cfpItem);

                //Get Unloads Bypass
                aTransactionItem = GetTransactionItem(CFP_UNLOADS_BYPASS);
                cfpItem = new CFPReconItem();
                cfpItem.Description = "UNLOADS BYPASS";
                if (aTransactionItem != null)
                {
                    //First set column DEBIT equals second set column CREDIT. First set column CREDIT equals second set column DEBIT.  
                    //Results on the second set of Debit and Credit column s/b the reverse of the first set of Debit and Credit column amount.
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        cfpItem.Credit2 = aTransactionItem.TransactionAmount;
                        mTotalLoadUnloadActivity += cfpItem.Debit1;
                        mTotalWireForToday += cfpItem.Credit2;
                    }
                    else if (aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        cfpItem.Debit2 = (aTransactionItem.TransactionAmount * -1);
                        mTotalLoadUnloadActivity += cfpItem.Credit1;
                        mTotalWireForToday += cfpItem.Debit2;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //VS4637 Vadd Loads Merchant Acq
                aTransactionItem = GetTransactionItem(CFP_LOADS_MERCHANT_ACQUIRER);
                cfpItem = new CFPReconItem();
                cfpItem.Description = "LOADS MERCHANT ACQUIRER";
                if (aTransactionItem != null)
                {
                    cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                    mTotalLoadUnloadActivity += cfpItem.Credit1;
                }
                CFPReconItemList.Add(cfpItem);

                //Get Reloads Merchant Acquirer
                aTransactionItem = GetTransactionItem(CFP_RELOADS_MERCHANT_ACQUIRER);
                cfpItem = new CFPReconItem();
                cfpItem.Description = "RELOADS MERCHANT ACQUIRER";
                if (aTransactionItem != null)
                {
                    cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                    mTotalLoadUnloadActivity += cfpItem.Credit1;
                }
                CFPReconItemList.Add(cfpItem);

                //VS4637 Vadd Unloads Merchant Acq
                aTransactionItem = GetTransactionItem(CFP_UNLOADS_MERCHANT_ACQUIRER);
                cfpItem = new CFPReconItem();
                cfpItem.Description = "UNLOADS MERCHANT ACQUIRER";
                if (aTransactionItem != null)
                {
                    cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                    mTotalLoadUnloadActivity += cfpItem.Credit1;
                }
                CFPReconItemList.Add(cfpItem);
#endregion

#region Total Load/Unload Activity

                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_TOTAL_LOAD_UNLOAD_ACTIVITY;
                //value is always positive...
                cfpItem.Credit1 = Math.Abs(mTotalLoadUnloadActivity);
                cfpReconItemList.Add(cfpItem);

                mTotalChangeToCFP += cfpItem.Credit1;

                //Get Transaction Disuptes
                aTransactionItem = GetTransactionItem(CFP_TRANSACTION_DISPUTES);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_TRANSACTION_DISPUTES;
                if (aTransactionItem != null)
                {
                    //Debit results in the trace file shall be entered as a positive number on the Credit side of the worksheet.  
                    //Credit results in the trace file shall be entered as a negative number on the Debit side of the worksheet .
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalChangeToCFP += cfpItem.Debit1;
                    }
                    else if (aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalChangeToCFP += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get Merchant Adjustments
                aTransactionItem = GetTransactionItem(CFP_MERCHANT_ADJUSTMENTS);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_MERCHANT_ADJUSTMENTS;
                if (aTransactionItem != null)
                {
                    //Debit results in the trace file shall be entered as a positive number on the Credit side of the worksheet.  
                    //Credit results in the trace file shall be entered as a negative number on the Debit side of the worksheet .
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalChangeToCFP += cfpItem.Debit1;
                    }
                    else if (aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalChangeToCFP += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get Manual Adjustments
                aTransactionItem = GetTransactionItem(CFP_MANUAL_ADJUSTMENTS);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_MANUAL_ADJUSTMENTS;
                if (aTransactionItem != null)
                {
                    //Debit results in the trace file shall be entered as a positive number on the Credit side of the worksheet.  
                    //Credit results in the trace file shall be entered as a negative number on the Debit side of the worksheet .
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalChangeToCFP += cfpItem.Debit1;
                    }
                    else if (aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalChangeToCFP += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //VS4649 Get Closed for Neg Bal
                aTransactionItem = GetTransactionItem(CFP_CLOSED_FOR_NEG_BAL);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_CLOSED_FOR_NEG_BAL;
                if (aTransactionItem != null)
                {
                    //Debit results in the trace file shall be entered as a positive number on the Credit side of the worksheet.  
                    //Credit results in the trace file shall be entered as a negative number on the Debit side of the worksheet .
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalChangeToCFP += cfpItem.Debit1;
                    }
                    else if (aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalChangeToCFP += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);


                //Get Closed for Esdcheatment
                aTransactionItem = GetTransactionItem(CFP_CLOSED_FOR_ESCHEATMENT);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_CLOSED_FOR_ESCHEATMENT;
                if (aTransactionItem != null)
                {
                    //Debit results in the trace file shall be entered as a positive number on the Credit side of the worksheet.  
                    //Credit results in the trace file shall be entered as a negative number on the Debit side of the worksheet .
                    if (aTransactionItem.DBCR1 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.TransactionAmount * -1);
                        mTotalChangeToCFP += cfpItem.Debit1;
                    }
                    else if (aTransactionItem.DBCR1 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.TransactionAmount;
                        mTotalChangeToCFP += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get Card PGM Fees
                aTransactionItem = GetTransactionItem(CFP_CARD_PGM_FEES);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_CARD_PGM_FEES;
                if (aTransactionItem != null)
                {
                    //Debit results in the trace file shall be entered as a positive number on the Credit side of the worksheet.  
                    //Credit results in the trace file shall be entered as a negative number on the Debit side of the worksheet .
                    if (aTransactionItem.DBCR2 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.FeeAmount * -1);
                        mTotalChangeFeeToIncome += cfpItem.Debit1;
                        mTotalChangeToCFP += cfpItem.Debit1;
                    }
                    else if (aTransactionItem.DBCR2 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.FeeAmount;
                        mTotalChangeFeeToIncome += cfpItem.Credit1;
                        mTotalChangeToCFP += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

                //Get Card PGM Fee Adjustment
                aTransactionItem = GetTransactionItem(CFP_CARD_PGM_FEE_ADJUSTMENTS);
                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_CARD_PGM_FEE_ADJUSTMENTS;
                if (aTransactionItem != null)
                {
                    //Debit results in the trace file shall be entered as a positive number on the Credit side of the worksheet.  
                    //Credit results in the trace file shall be entered as a negative number on the Debit side of the worksheet .
                    if (aTransactionItem.DBCR2 == "DR")
                    {
                        cfpItem.Debit1 = (aTransactionItem.FeeAmount * -1);
                        mTotalChangeFeeToIncome += cfpItem.Debit1;
                        mTotalChangeToCFP += cfpItem.Debit1;
                    }
                    else if (aTransactionItem.DBCR2 == "CR")
                    {
                        cfpItem.Credit1 = aTransactionItem.FeeAmount;
                        mTotalChangeFeeToIncome += cfpItem.Credit1;
                        mTotalChangeToCFP += cfpItem.Credit1;
                    }
                    else
                    {
                        //leave value empty for all the columns...
                    }
                }
                CFPReconItemList.Add(cfpItem);

#endregion

#region Total Change To Fee Income Account

                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_TOTAL_CHANGE_TO_FEE_INCOME_ACCOUNT;
                //value is flipped via a negative sign...
                cfpItem.Debit1 = (mTotalChangeFeeToIncome * -1);
                cfpReconItemList.Add(cfpItem);

#endregion

#region Total Wire For Today

                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_TOTAL_WIRE_FOR_TODAY;
                cfpItem.Credit2 = mTotalWireForToday;
                CFPReconItemList.Add(cfpItem);
#endregion

#region Total Change To CFP

                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_TOTAL_CHANGE_TO_CFP;
                cfpItem.Debit1 = mTotalChangeToCFP;
                CFPReconItemList.Add(cfpItem);
                mEndOfDayBalance += mTotalChangeToCFP;

#endregion

#region End Of Day Balance

                cfpItem = new CFPReconItem();
                cfpItem.Description = CFP_END_OF_DAY_BALANCE;
                cfpItem.Credit1 = mEndOfDayBalance;
                CFPReconItemList.Add(cfpItem);
#endregion
            }

            //OK finished massaging the data.  Now we need to assign the data over to Ultragrid so that we can use it to export to Excel...
            ultragridFinancialSummary.BindingContext = new System.Windows.Forms.BindingContext();
            UltraGridFinancialSummary.DataSource = CFPReconItemList;
            SetColumnHeaderLayout();
        }

        /// <summary>
        /// GetTransactionItem is a helper function that searches through the TransactionType list stored in the
        /// Global Summary Branch object looking for a transaction string that matches the description passed into the
        /// parameter value.  If nothing is found then a null value is returned.
        /// </summary>
        /// <param name="description">string name of the transaction</param>
        /// <returns></returns>
        private TransactionType GetTransactionItem(string description)
        {
            foreach (TransactionType aTransaction in GSummaryBranch.TransactionList)
            {
                if (aTransaction.Transaction.StartsWith(description, StringComparison.CurrentCultureIgnoreCase))
                    return aTransaction;
            }

            return null;
        }

        private FinancialSummaryItem GetFinancialSummaryItem(string description)
        {
            foreach (FinancialSummaryItem aItem in financialSummaryItemList)
            {
                if (aItem.Description.StartsWith(description, StringComparison.CurrentCultureIgnoreCase))
                    return aItem;
            }

            return null;
        }


#region end of report event

        public delegate void EndofReportEventHandler();
        public event EndofReportEventHandler EndofReport;

#endregion

    }
#region CFPReconItem class
    public class CFPReconItem
    {
        private string mdescription = "";
        public string Description
        {
            get { return mdescription; }

            set
            {
                if (!mdescription.Equals(value))
                {
                    mdescription = string.IsNullOrEmpty(value) ? "" : value.Trim();
                }
            }
        }        
        
        private decimal mdb1 = 0;
        public decimal Debit1
        {
            get { return mdb1; }
            set
            {
                if (mdb1 != value)
                {
                    mdb1 = value;
                }
            }
        }

        private decimal mcr1 = 0;
        public decimal Credit1
        {
            get { return mcr1; }
            set
            {
                if (mcr1 != value)
                {
                    mcr1 = value;
                }
            }
        }

        private decimal mdb2 = 0;
        public decimal Debit2
        {
            get { return mdb2; }
            set
            {
                if (mdb2 != value)
                {
                    mdb2 = value;
                }
            }
        }

        private decimal mcr2 = 0;
        public decimal Credit2
        {
            get { return mcr2; }
            set
            {
                if (mcr2 != value)
                {
                    mcr2 = value;
                }
            }
        }
    }
#endregion

#region FinancialSummaryItem class
    public class FinancialSummaryItem
    {

        private string mdescription = "";
        public string Description
        {
            get { return mdescription; }
            set
            {
                if (!mdescription.Equals(value))
                {
                    mdescription = string.IsNullOrEmpty(value) ? "" : value.Trim();
                }
            }
        }

        private decimal mdb = 0;
        public decimal Debit
        {
            get { return mdb; }
            set
            {
                if (mdb != value)
                {
                    mdb = value;
                }
            }
        }

        private decimal mcr = 0;
        public decimal Credit
        {
            get { return mcr; }
            set
            {
                if (mcr != value)
                {
                    mcr = value;
                }
            }
        }

        private decimal mnet = 0;
        public decimal Net
        {
            get { return mnet; }
            set
            {
                if (mnet != value)
                {
                    mnet = value;
                }
            }
        }

        private string mreportID = "";
        public string ReportID
        {
            get { return mreportID; }
            set
            {
                if (!mreportID.Equals(value))
                {
                    mreportID = string.IsNullOrEmpty(value) ? "" : value.Trim();
                }
            }
        }

        private string mDBCR = "";
        public string DBCR
        {
            get { return mDBCR; }
            set
            {
                mDBCR = value;
            }
        }
    }
#endregion
}

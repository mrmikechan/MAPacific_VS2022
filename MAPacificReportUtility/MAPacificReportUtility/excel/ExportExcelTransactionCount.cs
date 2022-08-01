using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using Infragistics.Win.UltraWinGrid;
using Infragistics.Win;

namespace MAPacificReportUtility.excel
{
    class ExportExcelTransactionCount
    {
        private List<TransactionCountItem> mtrnctlist;
        public List<TransactionCountItem> TransactionCountList
        {
            get { return mtrnctlist; }
            set
            {
                if (value != null)
                {
                    mtrnctlist = value;
                }
            }
        }

        private TransactionCountItem mTotalItem;
        public TransactionCountItem TransactionTotalItem
        {
            get { return mTotalItem; }
        }

        private UltraGrid ultragridTCount;
        public UltraGrid UltgragridTCount
        {
            get { return ultragridTCount; }
        }

        //decimal objects used to hold running total column data...
        private decimal mTotal_Loads_Fi_Funds_Trnsfr;
        private decimal mTotal_Reloads_Fi_Funds_Trnsfr;
        private decimal mTotal_Reloads_Bypass;
        private decimal mTotal_Reloads_Merchant_ACQ;
        private decimal mTotal_Unloads_Fi_Funds_Trnsfr;
        private decimal mTotal_Manual_Adjustments;
        private decimal mTotal_Total_Transaction_Cnt;
        private decimal mTotal_Unloads_Bypass;
        private decimal mTotal_Loads_Merch_Pos_Funding;
        //VS6160 Add ACH DIRECT DEPOSIT and RELOADS/MONEY TSFR RCVD
        private decimal mTotal_ACH_Direct_Deposit;
        private decimal mTotal_Reloads_Money_tsfr_rcvd;
        //VS4637 Add Loads Bypass, Loads Merch Acq, and Unloads Merch ACQ to this report
        private decimal mTotal_Loads_Bypass;
        private decimal mTotal_Loads_Merchant_ACQ;
        private decimal mTotal_Unloads_Merchant_ACQ;

        private bool isManAdjustTotalNegative = false;

        /// <summary>
        /// ExportExcelTransationAmount constructor.
        /// </summary>
        /// <param name="prepaidTotalSummary">BindingList of type ClientBranch that contains data for the Pre Paid Cards Total Summary</param>
        /// <param name="isBeken">bool isBeken used to determine if report is Beken or Map report type</param>
        public ExportExcelTransactionCount(System.ComponentModel.BindingList<ClientBranch> prepaidTotalSummary, bool isBeken)
        {
            TransactionCountList = new List<TransactionCountItem>();
            mTotal_Loads_Fi_Funds_Trnsfr = 0;
            mTotal_Reloads_Fi_Funds_Trnsfr = 0;
            mTotal_Reloads_Bypass = 0;
            mTotal_Reloads_Merchant_ACQ = 0;
            mTotal_Unloads_Fi_Funds_Trnsfr = 0;
            mTotal_Manual_Adjustments = 0;
            mTotal_Total_Transaction_Cnt = 0;
            mTotal_Loads_Merch_Pos_Funding = 0;
            mTotal_Unloads_Bypass = 0;
            //VS4637
            mTotal_Loads_Bypass = 0;
            mTotal_Loads_Merchant_ACQ = 0;
            mTotal_Unloads_Merchant_ACQ = 0;
            //vs6160
            mTotal_ACH_Direct_Deposit = 0;
            mTotal_Reloads_Money_tsfr_rcvd = 0;

            ultragridTCount = new UltraGrid();
            ultragridTCount.InitializeRow += new InitializeRowEventHandler(ultragridTCount_InitializeRow);
            ultragridTCount.InitializeLayout += new InitializeLayoutEventHandler(ultragridTCount_InitializeLayout);

            //Set the font to Arial.
            FontData fd1 = ultragridTCount.DisplayLayout.Appearance.FontData;
            fd1.Name = "Arial";

            ultragridTCount.BindingContext = new System.Windows.Forms.BindingContext();
            ProcessList(prepaidTotalSummary, isBeken);
            SetColumnHeaderLayout();
        }

        /// <summary>
        /// InitializeLayout event setup up the text alignment in the cells for each column when exported out to Excel.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultragridTCount_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            //VS6160:
            //Add ACH Direct Deposit after Reloads Merchant Acq
            //Add Reloads Money Tsfr Rcvd after ACH Direct Depost

            e.Layout.Bands[0].Columns["CLIENT_ID"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["CREDIT_UNION"].PerformAutoResize(Infragistics.Win.UltraWinGrid.PerformAutoSizeType.AllRowsInBand);
            e.Layout.Bands[0].Columns["LOADS_FI_FUNDS_TRNSFER"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["RELOAD_FI_FUNDS_TRNSFER"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["LOADS_MERCH_POS_FUNDING"].CellAppearance.TextHAlign = HAlign.Center;
            //VS4637 add Loads_Bypass
            e.Layout.Bands[0].Columns["LOADS_BYPASS"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["RELOADS_BYPASS"].CellAppearance.TextHAlign = HAlign.Center;
            //VS4637 add Loads_Merchant_ACQ
            e.Layout.Bands[0].Columns["LOADS_MERCHANT_ACQ"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["RELOADS_MERCHANT_ACQ"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["ACH_DIRECT_DEPOSIT"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["RELOADS_MONEY_TSFR_RCVD"].CellAppearance.TextHAlign = HAlign.Center;
            //VS4637 add Unloads_Merchant_ACQ
            e.Layout.Bands[0].Columns["UNLOADS_MERCHANT_ACQ"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["UNLOADS_FI_FUNDS_TRNSFR"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["UNLOADS_BYPASS"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["MANUAL_ADJUSTMENTS"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["TOTAL_TRANSACTION_COUNT"].CellAppearance.TextHAlign = HAlign.Center;
        }

        public void SetColumnHeaderLayout()
        {
            //Initialize the header columns in the Ultragrid to a specific font, bold, height, and color.  These settings
            //from the column header will then be used when exported into a Excel worksheet.
            foreach (UltraGridColumn col in ultragridTCount.DisplayLayout.Bands[0].Columns)
            {
                col.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                col.Header.Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                col.Header.Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                col.Header.Appearance.BackGradientStyle = GradientStyle.Vertical;
                col.Header.Appearance.FontData.Bold = DefaultableBoolean.True;
                col.Header.Appearance.FontData.SizeInPoints = 10;
                col.Header.Appearance.FontData.Name = "Arial";
                col.Header.Appearance.ForeColor = System.Drawing.Color.White;
                col.Header.Appearance.TextHAlign = HAlign.Center;
            }
        }


        void ultragridTCount_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            //Highlight the Total row in the grid so that when exported out to Excel that row is also highlighted.
            if (e.Row.Cells["CLIENT_ID"].Value.ToString().Equals("Total", StringComparison.CurrentCultureIgnoreCase))
            {
                e.Row.Cells["CLIENT_ID"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["CLIENT_ID"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["CLIENT_ID"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["CLIENT_ID"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["CLIENT_ID"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["CLIENT_ID"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["CLIENT_ID"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["CREDIT_UNION"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["CREDIT_UNION"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["CREDIT_UNION"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["CREDIT_UNION"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["CREDIT_UNION"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["CREDIT_UNION"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["CREDIT_UNION"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["LOADS_FI_FUNDS_TRNSFER"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["LOADS_FI_FUNDS_TRNSFER"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["LOADS_FI_FUNDS_TRNSFER"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["LOADS_FI_FUNDS_TRNSFER"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["LOADS_FI_FUNDS_TRNSFER"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["LOADS_FI_FUNDS_TRNSFER"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["LOADS_FI_FUNDS_TRNSFER"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["RELOAD_FI_FUNDS_TRNSFER"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["RELOAD_FI_FUNDS_TRNSFER"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["RELOAD_FI_FUNDS_TRNSFER"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["RELOAD_FI_FUNDS_TRNSFER"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["RELOAD_FI_FUNDS_TRNSFER"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["RELOAD_FI_FUNDS_TRNSFER"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["RELOAD_FI_FUNDS_TRNSFER"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["LOADS_MERCH_POS_FUNDING"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["LOADS_MERCH_POS_FUNDING"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["LOADS_MERCH_POS_FUNDING"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["LOADS_MERCH_POS_FUNDING"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["LOADS_MERCH_POS_FUNDING"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["LOADS_MERCH_POS_FUNDING"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["LOADS_MERCH_POS_FUNDING"].Appearance.ForeColor = System.Drawing.Color.White;

                //VS4637 Add Loads Bypass
                e.Row.Cells["LOADS_BYPASS"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["LOADS_BYPASS"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["LOADS_BYPASS"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["LOADS_BYPASS"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["LOADS_BYPASS"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["LOADS_BYPASS"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["LOADS_BYPASS"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["RELOADS_BYPASS"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["RELOADS_BYPASS"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["RELOADS_BYPASS"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["RELOADS_BYPASS"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["RELOADS_BYPASS"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["RELOADS_BYPASS"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["RELOADS_BYPASS"].Appearance.ForeColor = System.Drawing.Color.White;

                //VS4637 Add Loads Merchant Acq
                e.Row.Cells["LOADS_MERCHANT_ACQ"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["LOADS_MERCHANT_ACQ"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["LOADS_MERCHANT_ACQ"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["LOADS_MERCHANT_ACQ"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["LOADS_MERCHANT_ACQ"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["LOADS_MERCHANT_ACQ"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["LOADS_MERCHANT_ACQ"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["RELOADS_MERCHANT_ACQ"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["RELOADS_MERCHANT_ACQ"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["RELOADS_MERCHANT_ACQ"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["RELOADS_MERCHANT_ACQ"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["RELOADS_MERCHANT_ACQ"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["RELOADS_MERCHANT_ACQ"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["RELOADS_MERCHANT_ACQ"].Appearance.ForeColor = System.Drawing.Color.White;

                //VS6160 ACH DIRECT DEPOSIT and RELOADS MONEY TSFR RCVD
                e.Row.Cells["ACH_DIRECT_DEPOSIT"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["ACH_DIRECT_DEPOSIT"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["ACH_DIRECT_DEPOSIT"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["ACH_DIRECT_DEPOSIT"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["ACH_DIRECT_DEPOSIT"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["ACH_DIRECT_DEPOSIT"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["ACH_DIRECT_DEPOSIT"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["RELOADS_MONEY_TSFR_RCVD"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["RELOADS_MONEY_TSFR_RCVD"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["RELOADS_MONEY_TSFR_RCVD"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["RELOADS_MONEY_TSFR_RCVD"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["RELOADS_MONEY_TSFR_RCVD"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["RELOADS_MONEY_TSFR_RCVD"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["RELOADS_MONEY_TSFR_RCVD"].Appearance.ForeColor = System.Drawing.Color.White;

                //VS4637Add Unloads Merchant Acq
                e.Row.Cells["UNLOADS_MERCHANT_ACQ"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["UNLOADS_MERCHANT_ACQ"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["UNLOADS_MERCHANT_ACQ"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["UNLOADS_MERCHANT_ACQ"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["UNLOADS_MERCHANT_ACQ"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["UNLOADS_MERCHANT_ACQ"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["UNLOADS_MERCHANT_ACQ"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["UNLOADS_FI_FUNDS_TRNSFR"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["UNLOADS_FI_FUNDS_TRNSFR"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["UNLOADS_FI_FUNDS_TRNSFR"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["UNLOADS_FI_FUNDS_TRNSFR"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["UNLOADS_FI_FUNDS_TRNSFR"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["UNLOADS_FI_FUNDS_TRNSFR"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["UNLOADS_FI_FUNDS_TRNSFR"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["UNLOADS_BYPASS"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["UNLOADS_BYPASS"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["UNLOADS_BYPASS"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["UNLOADS_BYPASS"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["UNLOADS_BYPASS"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["UNLOADS_BYPASS"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["UNLOADS_BYPASS"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["MANUAL_ADJUSTMENTS"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["MANUAL_ADJUSTMENTS"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["MANUAL_ADJUSTMENTS"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["MANUAL_ADJUSTMENTS"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["MANUAL_ADJUSTMENTS"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["MANUAL_ADJUSTMENTS"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["MANUAL_ADJUSTMENTS"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["TOTAL_TRANSACTION_COUNT"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["TOTAL_TRANSACTION_COUNT"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["TOTAL_TRANSACTION_COUNT"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["TOTAL_TRANSACTION_COUNT"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["TOTAL_TRANSACTION_COUNT"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["TOTAL_TRANSACTION_COUNT"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["TOTAL_TRANSACTION_COUNT"].Appearance.ForeColor = System.Drawing.Color.White;
            }
        }

        /// <summary>
        /// ProcessList takes the BindingList of type ClientBranch and processes the data in there and massage them into a new List that
        /// is used to generate a new excel worksheet for Transaction Amount
        /// </summary>
        /// <param name="inList">List of ClientBranch types that contains the data needed to be processed</param>
        /// <param name="isBeken">bool isBeken used to determine if report is Beken or Map report type</param>
        private void ProcessList(System.ComponentModel.BindingList<ClientBranch> inList, bool isBeken)
        {

            TransactionCountItem newItem;
            foreach (ClientBranch item in inList)
            {
                if (isBeken)
                {
                    //Beken report type. We only want to see client ID related to Beken reports.
                    if (!item.Report.Equals(ClientBranch.ReportType.BEKEN))
                        continue;
                }

                if (!isBeken)
                {
                    if (item.Report.Equals(ClientBranch.ReportType.BEKEN))
                        continue;
                }

                newItem = new TransactionCountItem();
                newItem.CLIENT_ID = item.ClientID;
                newItem.CREDIT_UNION = item.Name;

                //look into the list of transactions in the ClientBranch
                //add the data into the new item and also calculate a running total value for each column...
                if (item.TransactionList.Count > 0)
                {
                    foreach (TransactionType aTransaction in item.TransactionList)
                    {
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.LOADS_FI_FUNDS_TRANSFER))
                        {
                            newItem.LOADS_FI_FUNDS_TRNSFER = aTransaction.TransactionCount;
                            mTotal_Loads_Fi_Funds_Trnsfr += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.RELOADS_FI_FUNDS_TRANSFER))
                        {
                            newItem.RELOAD_FI_FUNDS_TRNSFER = aTransaction.TransactionCount;
                            mTotal_Reloads_Fi_Funds_Trnsfr += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.LOADS_MERCH_POS_FUNDING))
                        {
                            newItem.LOADS_MERCH_POS_FUNDING = aTransaction.TransactionCount;
                            mTotal_Loads_Merch_Pos_Funding += aTransaction.TransactionCount;
                        }

                        //VS4637 Add Loads Bypass
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.LOADS_BYPASS))
                        {
                            newItem.LOADS_BYPASS = aTransaction.TransactionCount;
                            mTotal_Loads_Bypass += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.RELOADS_BYPASS))
                        {
                            newItem.RELOADS_BYPASS = aTransaction.TransactionCount;
                            mTotal_Reloads_Bypass += aTransaction.TransactionCount;
                        }

                        //VS4637 Add Loads Merchant Acq
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.LOADS_MERCHANT_ACQ))
                        {
                            newItem.LOADS_MERCHANT_ACQ = aTransaction.TransactionCount;
                            mTotal_Loads_Merchant_ACQ += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.RELOADS_MERCHANT_ACQ))
                        {
                            newItem.RELOADS_MERCHANT_ACQ = aTransaction.TransactionCount;
                            mTotal_Reloads_Merchant_ACQ += aTransaction.TransactionCount;
                        }

                        //VS4637 Add Unloads Merchant Acq
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.UNLOADS_MERCHANT_ACQ))
                        {
                            newItem.UNLOADS_MERCHANT_ACQ = aTransaction.TransactionCount;
                            mTotal_Unloads_Merchant_ACQ += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.UNLOADS_FI_FUNDS_TRANSFER))
                        {
                            newItem.UNLOADS_FI_FUNDS_TRNSFR = aTransaction.TransactionCount;
                            mTotal_Unloads_Fi_Funds_Trnsfr += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.UNLOADS_BYPASS))
                        {
                            newItem.UNLOADS_BYPASS = aTransaction.TransactionCount;
                            mTotal_Unloads_Bypass += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.MANUAL_ADJUSTMENT))
                        {
                            newItem.MANUAL_ADJUSTMENTS = aTransaction.TransactionCount;
                            mTotal_Manual_Adjustments += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.TOTAL_LOAD_UNLOAD_ACTIVITY))
                        {
                            newItem.TOTAL_TRANSACTION_COUNT = aTransaction.TransactionCount;
                            mTotal_Total_Transaction_Cnt += aTransaction.TransactionCount;
                        }

                        //VS6160 Add ACH Direct Deposit
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.ACH_DIRECT_DEPOSIT))
                        {

                            newItem.ACH_DIRECT_DEPOSIT = aTransaction.TransactionCount;
                            mTotal_ACH_Direct_Deposit += aTransaction.TransactionCount;
                        }

                        //VS6160 Add Reloads Money Tsfr Rcvd
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.RELOADS_MONEY_TRSFR_RCVD))
                        {
                            newItem.RELOADS_MONEY_TSFR_RCVD = aTransaction.TransactionCount;
                            mTotal_Reloads_Money_tsfr_rcvd += aTransaction.TransactionCount;
                        }
                    }
                }

                TransactionCountList.Add(newItem);
            }
            //finished processing the BindingList... now lets create a new row to add the total columns into the new list...
            mTotalItem = new TransactionCountItem();
            mTotalItem.CLIENT_ID = "Total";
            mTotalItem.CREDIT_UNION = "";
            mTotalItem.LOADS_FI_FUNDS_TRNSFER = mTotal_Loads_Fi_Funds_Trnsfr;
            mTotalItem.MANUAL_ADJUSTMENTS = mTotal_Manual_Adjustments;
            mTotalItem.RELOAD_FI_FUNDS_TRNSFER = mTotal_Reloads_Fi_Funds_Trnsfr;
            mTotalItem.RELOADS_BYPASS = mTotal_Reloads_Bypass;
            mTotalItem.RELOADS_MERCHANT_ACQ = mTotal_Reloads_Merchant_ACQ;
            mTotalItem.TOTAL_TRANSACTION_COUNT = mTotal_Total_Transaction_Cnt;
            mTotalItem.UNLOADS_FI_FUNDS_TRNSFR = mTotal_Unloads_Fi_Funds_Trnsfr;
            mTotalItem.LOADS_MERCH_POS_FUNDING = mTotal_Loads_Merch_Pos_Funding;
            mTotalItem.UNLOADS_BYPASS = mTotal_Unloads_Bypass;
            //VS4637 Add the values from Load Bypass, Loads Merchant Acq, and Unloads Merchant Acq
            mTotalItem.LOADS_BYPASS = mTotal_Loads_Bypass;
            mTotalItem.LOADS_MERCHANT_ACQ = mTotal_Loads_Merchant_ACQ;
            mTotalItem.UNLOADS_MERCHANT_ACQ = mTotal_Unloads_Merchant_ACQ;
            //VS6160 Add ACH Direct Deposit and Reloads Money Tsfr Rcvd
            mTotalItem.ACH_DIRECT_DEPOSIT = mTotal_ACH_Direct_Deposit;
            mTotalItem.RELOADS_MONEY_TSFR_RCVD = mTotal_Reloads_Money_tsfr_rcvd;

            TransactionCountList.Add(mTotalItem);

            ultragridTCount.DataSource = TransactionCountList;
        }
    }

    //Note that the order of properties listed is important.
    //They are displayed in this specific order when exported into Excel.
    class TransactionCountItem : INotifyPropertyChanged
    {

        private string mclientID = "";
        public string CLIENT_ID
        {
            get { return mclientID; }
            set
            {
                if (!mclientID.Equals(value))
                {
                    mclientID = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("CLIENT_ID");
                }
            }
        }

        private string mcreditUnion = "";
        public string CREDIT_UNION
        {
            get { return mcreditUnion; }
            set
            {
                if (!mcreditUnion.Equals(value))
                {
                    mcreditUnion = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("CREDIT_UNION");
                }
            }
        }

        private decimal mloadsfifundstrnsfr = 0;
        public decimal LOADS_FI_FUNDS_TRNSFER
        {
            get { return mloadsfifundstrnsfr; }
            set
            {
                if (mloadsfifundstrnsfr != value)
                {
                    mloadsfifundstrnsfr = value;
                    OnPropertyChanged("LOADS_FI_FUNDS_TRNSFER");
                }
            }
        }

        private decimal mreloadsfifundstrnsfr = 0;
        public decimal RELOAD_FI_FUNDS_TRNSFER
        {
            get { return mreloadsfifundstrnsfr; }
            set
            {
                if (mreloadsfifundstrnsfr != value)
                {
                    mreloadsfifundstrnsfr = value;
                    OnPropertyChanged("RELOAD_FI_FUNDS_TRNSFER");
                }
            }
        }

        private decimal mloadsmerchposfunding = 0;
        public decimal LOADS_MERCH_POS_FUNDING
        {
            get { return mloadsmerchposfunding; }
            set
            {
                if (mloadsmerchposfunding != value)
                {
                    mloadsmerchposfunding = value;
                    OnPropertyChanged("LOADS_MERCH_POS_FUNDING");
                }
            }
        }

        //VS4637
        private decimal mloadsbypass = 0;
        public decimal LOADS_BYPASS
        {
            get { return mloadsbypass; }
            set
            {
                if (mloadsbypass != value)
                {
                    mloadsbypass = value;
                    OnPropertyChanged("LOADS_BYPASS");
                }
            }
        }

        private decimal mreloadsbypass = 0;
        public decimal RELOADS_BYPASS
        {
            get { return mreloadsbypass; }
            set
            {
                if (mreloadsbypass != value)
                {
                    mreloadsbypass = value;
                    OnPropertyChanged("RELOADS_BYPASS");
                }
            }
        }

        //VS4637
        private decimal mloadsmerchanacq = 0;
        public decimal LOADS_MERCHANT_ACQ
        {
            get { return mloadsmerchanacq; }
            set
            {
                if (mloadsmerchanacq != value)
                {
                    mloadsmerchanacq = value;
                    OnPropertyChanged("LOADS_MERCHANT_ACQ");
                }
            }
        }
        private decimal mreloadsmerchantacq = 0;
        public decimal RELOADS_MERCHANT_ACQ
        {
            get { return mreloadsmerchantacq; }
            set
            {
                if (mreloadsmerchantacq != value)
                {
                    mreloadsmerchantacq = value;
                    OnPropertyChanged("RELOADS_MERCHANT_ACQ");
                }
            }
        }

        //VS6160 Add ACH Direct Deposit
        private decimal machdirectdeposit = 0;
        public decimal ACH_DIRECT_DEPOSIT
        {
            get { return machdirectdeposit; }
            set
            {
                if (machdirectdeposit != value)
                {
                    machdirectdeposit = value;
                    OnPropertyChanged("ACH_DIRECT_DEPOSIT");
                }
            }
        }

        //VS6160 Add Reloads Money Tsfr Rcvd
        private decimal mreloadsmoneytsfrcvd = 0;
        public decimal RELOADS_MONEY_TSFR_RCVD
        {
            get { return mreloadsmoneytsfrcvd; }
            set
            {
                if (mreloadsmoneytsfrcvd != value)
                {
                    mreloadsmoneytsfrcvd = value;
                    OnPropertyChanged("RELOADS_MONEY_TSFR_RCVD");
                }
            }
        }

        //VS4637
        private decimal munloadsmerchanacq = 0;
        public decimal UNLOADS_MERCHANT_ACQ
        {
            get { return munloadsmerchanacq; }
            set
            {
                if (munloadsmerchanacq != value)
                {
                    mloadsmerchanacq = value;
                    OnPropertyChanged("UNLOADS_MERCHANT_ACQ");
                }
            }
        }

        private decimal munloadsfifundstrnsfr = 0;
        public decimal UNLOADS_FI_FUNDS_TRNSFR
        {
            get { return munloadsfifundstrnsfr; }
            set
            {
                if (munloadsfifundstrnsfr != value)
                {
                    munloadsfifundstrnsfr = value;
                    OnPropertyChanged("UNLOADS_FI_FUNDS_TRNSFR");
                }
            }
        }

        private decimal munloadsbypass = 0;
        public decimal UNLOADS_BYPASS
        {
            get { return munloadsbypass; }
            set
            {
                if (munloadsbypass != value)
                {
                    munloadsbypass = value;
                    OnPropertyChanged("UNLOADS_BYPASS");
                }
            }
        }

        private decimal mtotaltransactioncnt = 0;
        public decimal TOTAL_TRANSACTION_COUNT
        {
            get { return mtotaltransactioncnt; }
            set
            {
                if (mtotaltransactioncnt != value)
                {
                    mtotaltransactioncnt = value;
                    OnPropertyChanged("TOTAL_TRANSACTION_COUNT");
                }
            }
        }


        private decimal mmanualadjustment = 0;
        public decimal MANUAL_ADJUSTMENTS
        {
            get { return mmanualadjustment; }
            set
            {
                if (mmanualadjustment != value)
                {
                    mmanualadjustment = value;
                    OnPropertyChanged("MANUAL_ADJUSTMENTS");
                }
            }
        }

        #region INotifyPropertyChanged Members

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new System.ComponentModel.PropertyChangedEventArgs(name));
            }
        }
        #endregion

    }
}
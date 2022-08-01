using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using Infragistics.Win.UltraWinGrid;
using Infragistics.Win;

namespace MAPacificReportUtility.excel
{
    class ExportExcelCardActivityAmount
    {
        private List<TransactionCardActivityAmountItem> mtrnamtlist;
        public List<TransactionCardActivityAmountItem> TransactionCardAmountList
        {
            get { return mtrnamtlist; }
            set
            {
                if (value != null)
                {
                    mtrnamtlist = value;
                }
            }
        }

        private UltraGrid ultragridCardAmount;
        public UltraGrid UltgragridCardAmount
        {
            get { return ultragridCardAmount; }
        }

        //decimal objects used to hold running total column data...
        private decimal mTotal_Load_Disputes;
        private decimal mTotal_Purchases_Quasi_Cash;
        private decimal mTotal_Purchases_With_Cash_Back;
        private decimal mTotal_AFT_AA_PP;
        private decimal mTotal_Purchase_Returns;
        private decimal mTotal_Manual_Cash;
        private decimal mTotal_Atm_Cash;
        private decimal mTotal_Balance_Inquiries;
        private decimal mTotal_Rewards;
        private decimal mTotal_Total_Card_Activity;


        /// <summary>
        /// ExportExcelTransationAmount constructor.
        /// </summary>
        /// <param name="prepaidTotalSummary">BindingList of type ClientBranch that contains data for the Pre Paid Cards Total Summary</param>
        /// <param name="isBeken">bool isBeken used to determine if report is Beken or Map report type</param>
        public ExportExcelCardActivityAmount(System.ComponentModel.BindingList<ClientBranch> prepaidTotalSummary, bool isBeken)
        {
            TransactionCardAmountList = new List<TransactionCardActivityAmountItem>();
            mTotal_Load_Disputes = 0;
            mTotal_Purchases_Quasi_Cash = 0;
            mTotal_Purchases_With_Cash_Back = 0;
            mTotal_AFT_AA_PP = 0;
            mTotal_Purchase_Returns = 0;
            mTotal_Manual_Cash = 0;
            mTotal_Atm_Cash = 0;
            mTotal_Balance_Inquiries = 0;
            mTotal_Rewards = 0;
            mTotal_Total_Card_Activity = 0;

            ultragridCardAmount = new UltraGrid();
            ultragridCardAmount.InitializeRow += new InitializeRowEventHandler(ultragridCardAmount_InitializeRow);
            ultragridCardAmount.InitializeLayout += new InitializeLayoutEventHandler(ultragridCardAmount_InitializeLayout);

            //Set the font datatype to Arial so that when we export the grid to Excel the font will also be Arial.
            FontData fd1 = ultragridCardAmount.DisplayLayout.Appearance.FontData;
            fd1.Name = "Arial";

            //Need to set the Binding Context first when creating ultragrid during runtime.
            ultragridCardAmount.BindingContext = new System.Windows.Forms.BindingContext();
            ProcessList(prepaidTotalSummary, isBeken);
            SetColumnHeaderLayout();
        }

        /// <summary>
        /// Use the InitializeLayout event to setup the Text alignment in each column to be horizontal center aligned.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultragridCardAmount_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Bands[0].Columns["CLIENT_ID"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["CREDIT_UNION"].PerformAutoResize(Infragistics.Win.UltraWinGrid.PerformAutoSizeType.AllRowsInBand);
            e.Layout.Bands[0].Columns["LOAD_DISPUTES"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["PURCHASES_QUASI_CASH"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["PURCHASES_WITH_CASH_BACK"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["AFT_AA_PP"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["PURCHASE_RETURNS"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["MANUAL_CASH"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["ATM_CASH"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["BALANCE_INQUIRIES"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["REWARDS"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["TOTAL_CARD_ACTIVITY"].CellAppearance.TextHAlign = HAlign.Center;

        }

        /// <summary>
        /// ultraGridTAmount_InitializeRow event used to determine when we hit the last row of data that contains the summary
        /// for all the columns.  Highlight this last row of cells.
        /// Note that this event is only used by ultraGridTAmount object.
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">InitializeRowEventArgs</param>
        void ultragridCardAmount_InitializeRow(object sender, InitializeRowEventArgs e)
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

                e.Row.Cells["LOAD_DISPUTES"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["LOAD_DISPUTES"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["LOAD_DISPUTES"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["LOAD_DISPUTES"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["LOAD_DISPUTES"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["LOAD_DISPUTES"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["LOAD_DISPUTES"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.ForeColor = System.Drawing.Color.White;


                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["AFT_AA_PP"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["AFT_AA_PP"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["AFT_AA_PP"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["AFT_AA_PP"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["AFT_AA_PP"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["AFT_AA_PP"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["AFT_AA_PP"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["PURCHASE_RETURNS"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["MANUAL_CASH"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["MANUAL_CASH"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["MANUAL_CASH"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["MANUAL_CASH"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["MANUAL_CASH"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["MANUAL_CASH"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["MANUAL_CASH"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["ATM_CASH"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["ATM_CASH"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["ATM_CASH"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["ATM_CASH"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["ATM_CASH"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["ATM_CASH"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["ATM_CASH"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["REWARDS"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["REWARDS"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["REWARDS"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["REWARDS"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["REWARDS"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["REWARDS"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["REWARDS"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.ForeColor = System.Drawing.Color.White;
            }
        }

        public void SetColumnHeaderLayout()
        {
            //Initialize the header columns in the Ultragrid to a specific font, bold, height, and color.  These settings
            //from the column header will then be used when exported into a Excel worksheet.
            foreach (UltraGridColumn col in ultragridCardAmount.DisplayLayout.Bands[0].Columns)
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

        /// <summary>
        /// ProcessList takes the BindingList of type ClientBranch and processes the data in there and massage them into a new List that
        /// is used to generate a new excel worksheet for Transaction Amount
        /// </summary>
        /// <param name="inList">List of ClientBranch types that contains the data needed to be processed</param>
        /// <param name="isBeken">bool isBeken used to determine if report is Beken or Map report type</param>
        private void ProcessList(System.ComponentModel.BindingList<ClientBranch> inList, bool isBeken)
        {

            TransactionCardActivityAmountItem newItem;
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

                newItem = new TransactionCardActivityAmountItem();
                newItem.CLIENT_ID = item.ClientID;
                newItem.CREDIT_UNION = item.Name;

                //look into the list of transactions in the ClientBranch
                //add the data into the new item and also calculate a running total value for each column...
                if (item.CardActivityTransactionList.Count > 0)
                {
                    foreach (TransactionType aTransaction in item.CardActivityTransactionList)
                    {
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.LOAD_DISPUTES))
                        {
                            if (aTransaction.DBCR1.Equals("CR", StringComparison.CurrentCultureIgnoreCase))
                            {
                                newItem.LOAD_DISPUTES = (aTransaction.TransactionAmount * -1); //make negative number value
                            }
                            else
                            {
                                newItem.LOAD_DISPUTES = aTransaction.TransactionAmount;
                            }
                            mTotal_Load_Disputes += newItem.LOAD_DISPUTES;

                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.PURCHASES_QUASI_CASH))
                        {
                            if (aTransaction.DBCR1.Equals("CR", StringComparison.CurrentCultureIgnoreCase))
                            {
                                newItem.PURCHASES_QUASI_CASH = (aTransaction.TransactionAmount * -1); //make negative number value
                            }
                            else
                            {
                                newItem.PURCHASES_QUASI_CASH = aTransaction.TransactionAmount;
                            }
                            mTotal_Purchases_Quasi_Cash += newItem.PURCHASES_QUASI_CASH;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.PURCHASES_WITH_CASH_BACK))
                        {
                            if (aTransaction.DBCR1.Equals("CR", StringComparison.CurrentCultureIgnoreCase))
                            {
                                newItem.PURCHASES_WITH_CASH_BACK = (aTransaction.TransactionAmount * -1); //make negative number value
                            }
                            else
                            {
                                newItem.PURCHASES_WITH_CASH_BACK = aTransaction.TransactionAmount;
                            }
                            mTotal_Purchases_With_Cash_Back += newItem.PURCHASES_WITH_CASH_BACK;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.AFT_AA_PP))
                        {

                            if (aTransaction.DBCR1.Equals("CR", StringComparison.CurrentCultureIgnoreCase))
                            {
                                newItem.AFT_AA_PP = (aTransaction.TransactionAmount * -1);
                            }
                            else
                            {
                                newItem.AFT_AA_PP = aTransaction.TransactionAmount;
                            }
                            mTotal_AFT_AA_PP += newItem.AFT_AA_PP;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.PURCHASE_RETURNS))
                        {
                            if (aTransaction.DBCR1.Equals("CR", StringComparison.CurrentCultureIgnoreCase))
                            {

                                newItem.PURCHASE_RETURNS = (aTransaction.TransactionAmount * -1);
                            }
                            else
                            {
                                newItem.PURCHASE_RETURNS = aTransaction.TransactionAmount;
                            }
                            mTotal_Purchase_Returns += newItem.PURCHASE_RETURNS;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.MANUAL_CASH))
                        {
                            if (aTransaction.DBCR1.Equals("CR", StringComparison.CurrentCultureIgnoreCase))
                            {

                                newItem.MANUAL_CASH = (aTransaction.TransactionAmount * -1);
                            }
                            else
                            {
                                newItem.MANUAL_CASH = aTransaction.TransactionAmount;
                            }
                            mTotal_Manual_Cash += newItem.MANUAL_CASH;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.ATM_CASH))
                        {
                            if (aTransaction.DBCR1.Equals("CR", StringComparison.CurrentCultureIgnoreCase))
                            {
                                newItem.ATM_CASH = (aTransaction.TransactionAmount * -1);
                            }
                            else
                            {
                                newItem.ATM_CASH = aTransaction.TransactionAmount;
                            }
                            mTotal_Atm_Cash += newItem.ATM_CASH;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.BALANCE_INQUIRIES))
                        {
                            if (aTransaction.DBCR1.Equals("CR", StringComparison.CurrentCultureIgnoreCase))
                            {
                                newItem.BALANCE_INQUIRIES = (aTransaction.TransactionAmount * -1);
                            }
                            else
                            {
                                newItem.BALANCE_INQUIRIES = aTransaction.TransactionAmount;
                            }
                            mTotal_Balance_Inquiries += newItem.BALANCE_INQUIRIES;
                        }

                        //VS4637 Add Unloads Merchant Acq
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.REWARDS))
                        {
                            if (aTransaction.DBCR1.Equals("CR", StringComparison.CurrentCultureIgnoreCase))
                            {
                                newItem.REWARDS = (aTransaction.TransactionAmount * -1);
                            }
                            else
                            {
                                newItem.REWARDS = aTransaction.TransactionAmount;
                            }
                            mTotal_Rewards += newItem.REWARDS;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.TOTAL_CARD_ACTIVITY))
                        {
                            if (aTransaction.DBCR1.Equals("CR", StringComparison.CurrentCultureIgnoreCase))
                            {
                                newItem.TOTAL_CARD_ACTIVITY = (aTransaction.TransactionAmount * -1);
                            }
                            else
                            {
                                newItem.TOTAL_CARD_ACTIVITY = aTransaction.TransactionAmount;
                            }
                            mTotal_Total_Card_Activity += newItem.TOTAL_CARD_ACTIVITY;
                        }
                    }
                }

                TransactionCardAmountList.Add(newItem);
            }
            //finished processing the BindingList... now lets create a new row to add the total columns into the new list...
            TransactionCardActivityAmountItem itemTotal = new TransactionCardActivityAmountItem();
            itemTotal.CLIENT_ID = "Total";
            itemTotal.CREDIT_UNION = "";
            itemTotal.LOAD_DISPUTES = mTotal_Load_Disputes;

            itemTotal.PURCHASES_QUASI_CASH = mTotal_Purchases_Quasi_Cash;
            //VS4068
            //Because we are overiding the font color and size in the Row for Total we have to
            //also accomodate the red color for negative value and remove the negative sign at the same time.  Normally this is
            //taken care of by the Excel format strings but that gets overriden. We also set the flag for isManAdjustTotalNegative value here. We do
            //not reset the value back to false in the constructor because for some odd reason the constructor gets called again after we have set the
            //value to true thus overiding it and we lose the flag to determine if we need to set the font color to red or not.
            //if (mTotal_Manual_Adjustments < 0)
            //{
            //    isManAdjustTotalNegative = true;
            //    //mTotal_Manual_Adjustments = Math.Abs(mTotal_Manual_Adjustments);
            //}
            //else
            //{
            //    isManAdjustTotalNegative = false;
            //}
            itemTotal.PURCHASES_WITH_CASH_BACK = mTotal_Purchases_With_Cash_Back;
            itemTotal.AFT_AA_PP = mTotal_AFT_AA_PP;
            itemTotal.PURCHASE_RETURNS = mTotal_Purchase_Returns;
            itemTotal.MANUAL_CASH = mTotal_Manual_Cash;
            itemTotal.ATM_CASH = mTotal_Atm_Cash;
            itemTotal.BALANCE_INQUIRIES = mTotal_Balance_Inquiries;
            itemTotal.REWARDS = mTotal_Rewards;
            itemTotal.TOTAL_CARD_ACTIVITY = mTotal_Total_Card_Activity;

            TransactionCardAmountList.Add(itemTotal);
            ultragridCardAmount.DataSource = TransactionCardAmountList;
        }
    }

    //Note that the order of properties listed is important.
    //They are displayed in this specific order when exported into Excel.
    class TransactionCardActivityAmountItem : INotifyPropertyChanged
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

        private decimal mloadsdisputes = 0;
        public decimal LOAD_DISPUTES
        {
            get { return mloadsdisputes; }
            set
            {
                if (mloadsdisputes != value)
                {
                    mloadsdisputes = value;
                    OnPropertyChanged("LOAD_DISPUTES");
                }
            }
        }

        private decimal mpurchasesquasicash = 0;
        public decimal PURCHASES_QUASI_CASH
        {
            get { return mpurchasesquasicash; }
            set
            {
                if (mpurchasesquasicash != value)
                {
                    mpurchasesquasicash = value;
                    OnPropertyChanged("PURCHASES_QUASI_CASH");
                }
            }
        }

        private decimal mpurchaseswithcashback = 0;
        public decimal PURCHASES_WITH_CASH_BACK
        {
            get { return mpurchaseswithcashback; }
            set
            {
                if (mpurchaseswithcashback != value)
                {
                    mpurchaseswithcashback = value;
                    OnPropertyChanged("PURCHASES_WITH_CASH_BACK");
                }
            }
        }

        private decimal maftaapp = 0;
        public decimal AFT_AA_PP
        {
            get { return maftaapp; }
            set
            {
                if (maftaapp != value)
                {
                    maftaapp = value;
                    OnPropertyChanged("AFT_AA_PP");
                }
            }
        }

        //VS4637
        private decimal mpurchasereturns = 0;
        public decimal PURCHASE_RETURNS
        {
            get { return mpurchasereturns; }
            set
            {
                if (mpurchasereturns != value)
                {
                    mpurchasereturns = value;
                    OnPropertyChanged("PURCHASE_RETURNS");
                }
            }
        }

        private decimal mmanualcash = 0;
        public decimal MANUAL_CASH
        {
            get { return mmanualcash; }
            set
            {
                if (mmanualcash != value)
                {
                    mmanualcash = value;
                    OnPropertyChanged("MANUAL_CASH");
                }
            }
        }

        //VS4637
        private decimal matmcash = 0;
        public decimal ATM_CASH
        {
            get { return matmcash; }
            set
            {
                if (matmcash != value)
                {
                    matmcash = value;
                    OnPropertyChanged("ATM_CASH");
                }
            }
        }
        private decimal mbalanceinquiries = 0;
        public decimal BALANCE_INQUIRIES
        {
            get { return mbalanceinquiries; }
            set
            {
                if (mbalanceinquiries != value)
                {
                    mbalanceinquiries = value;
                    OnPropertyChanged("BALANCE_INQUIRIES");
                }
            }
        }

        //VS4637
        private decimal mrewards = 0;
        public decimal REWARDS
        {
            get { return mrewards; }
            set
            {
                if (mrewards != value)
                {
                    matmcash = value;
                    OnPropertyChanged("REWARDS");
                }
            }
        }

        private decimal mtotalcardactivity = 0;
        public decimal TOTAL_CARD_ACTIVITY
        {
            get { return mtotalcardactivity; }
            set
            {
                if (mtotalcardactivity != value)
                {
                    mtotalcardactivity = value;
                    OnPropertyChanged("TOTAL_CARD_ACTIVITY");
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

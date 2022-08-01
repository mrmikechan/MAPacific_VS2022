using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace MAPacificReportUtility
{
    public class TransactionType : INotifyPropertyChanged
    {
        public TransactionType()
        {

        }

        #region data members
        public struct TransactionOption
        {
            public static string LOADS_FI_FUNDS_TRANSFER
            {
                get { return "LOADS/FI FUNDS TRANSFER"; }
            }

            public static string RELOADS_FI_FUNDS_TRANSFER
            {
                get { return "RELOADS/FI FUNDS TRNSFER"; }
            }

            public static string UNLOADS_FI_FUNDS_TRANSFER
            {
                get { return "UNLOADS/FI FUNDS TRNSFER"; }
            }
            //VS4637 Add LOADS_BYPASS per Orlando
            public static string LOADS_BYPASS
            {
                get{ return "LOADS/BY_PASS";}
            }

            //VS3594 Add unloads_bypass per Nikolay...
            public static string UNLOADS_BYPASS
            {
                get { return "UNLOADS/BY_PASS"; }
            }

            public static string RELOADS_BYPASS
            {
                get { return "RELOADS/BY-PASS"; }
            }

            //VS4637 Add LOADS_MERCHANT_ACQ per Orlando
            public static string LOADS_MERCHANT_ACQ
            {
                get { return "LOADS/MERCHANT ACQUIRER"; }
            }

            public static string RELOADS_MERCHANT_ACQ
            {
                get { return "RELOADS/MERCHANT ACQ"; }
            }

            public static string UNLOADS_MERCHANT_ACQ
            {
                get { return "UNLOADS/MERCHANT ACQ"; }
            }

            public static string MANUAL_ADJUSTMENT
            {
                get { return "MANUAL ADJUSTMENTS"; }
            }

            public static string TOTAL_LOAD_UNLOAD_ACTIVITY
            {
                get { return "*TOTAL LOAD/UNLOAD ACTIVITY"; }
            }

            //VS4066 Add per Orlando
            public static string LOADS_MERCH_POS_FUNDING
            {
                get { return "LOADS/MERCH POS FUNDING"; }
            }

            //VS4731 and VS4732 adding the Card Activity Specific transaction type Note that these transactions fall into Total Card Activity

            public static string LOAD_DISPUTES
            {

                get { return "LOAD DISPUTES"; }
            }

            public static string PURCHASES_QUASI_CASH
            {
                get { return "PURCHASES/QUASI CASH"; }
            }

            public static string PURCHASES_WITH_CASH_BACK
            {
                get { return "PURCHASES WITH CASH BACK"; }
            }

            public static string AFT_AA_PP
            {
                get { return "AFT - AA/PP"; }
            }

            public static string PURCHASE_RETURNS
            {
                get { return "PURCHASE RETURNS"; }
            }

            public static string MANUAL_CASH
            {
                get { return "MANUAL CASH"; }
            }

            public static string ATM_CASH
            {
                get { return "ATM CASH"; }
            }

            public static string BALANCE_INQUIRIES
            {
                get { return "BALANCE INQUIRIES"; }
            }
            
            public static string REWARDS
            {
                get { return "REWARDS"; }
            }
            
            public static string TOTAL_CARD_ACTIVITY
            {
                get { return "*TOTAL CARD ACTIVITY"; }
            }
            //end VS4731 and 4732

            //VS6159 Add new transaction type to process
            public static string ACH_DIRECT_DEPOSIT
            {
                get { return "ACH DIRECT DEPOSIT"; }
            }

            public static string RELOADS_MONEY_TRSFR_RCVD
            {
                get { return "RELOADS/MONEY TSFR RCVD"; }
            }
        }

        private string _transactioninfo = "";
        /// <summary>
        /// Transaction detail can be the following values (Unless used by GlobalSummaryBranch then it can be anything):
        /// Loads/FI Funds Transfer
        /// Reloads/FI Funds Transfer
        /// Unloads/FI Funds Transfer
        /// Reloads/ByPass
        /// Reloads/Merchant ACQ
        /// Unloads/Merchant ACQ
        /// Manual Adjustments
        /// Total Load/Unload Activity
        /// </summary>
        public string Transaction
        {
            get { return _transactioninfo; }
            set
            {
                if (!_transactioninfo.Equals(value))
                {
                    _transactioninfo = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("Transaction");
                }
            }
        }

        private decimal _trancount;

        public decimal TransactionCount
        {
            get { return _trancount; }
            set
            {
                if (_trancount != value)
                {
                    _trancount = value;
                    OnPropertyChanged("TransactionCount");
                }
            }
        }

        private decimal _tranamount;

        public decimal TransactionAmount
        {
            get { return _tranamount; }
            set
            {
                if (_tranamount!= value)
                {
                    _tranamount = value;
                    if (_tranamount == 0)
                    {
                        DBCR1 = "";
                    }
                    OnPropertyChanged("TransactionAmount");
                }
            }
        }

        private string _dbcr1 = "";
        public string DBCR1
        {
            get { return _dbcr1; }
            set
            {
                if (!_dbcr1.Equals(value))
                {
                    _dbcr1 = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("DBCR1");
                }
            }
        }

        private decimal _feeamount;
        public decimal FeeAmount
        {
            get { return _feeamount; }
            set
            {
                if (_feeamount!=value)
                {
                    _feeamount = value;
                    if (_feeamount == 0)
                    {
                        DBCR2 = "";
                    }
                    OnPropertyChanged("FeeAmount");
                }
            }
        }

        private string _dbcr2 = "";
        public string DBCR2
        {
            get { return _dbcr2; }
            set
            {
                if (!_dbcr2.Equals(value))
                {
                    _dbcr2 = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("DBCR2");
                }
            }
        }

        private decimal _totalamount;
        public decimal TotalAmount
        {
            get { return _totalamount; }
            set
            {
                if (_totalamount!=value)
                {
                    _totalamount = value;
                    if (_totalamount == 0)
                    {
                        DBCR3 = "";
                    }
                    OnPropertyChanged("TotalAmount");
                }
            }
        }

        private string _dbcr3 = "";
        public string DBCR3
        {
            get { return _dbcr3; }
            set
            {
                if (!_dbcr3.Equals(value))
                {
                    _dbcr3 = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("DBCR3");
                }
            }
        }

        #endregion


        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }
        #endregion

        #region utility function to clone object...

        public TransactionType Clone()
        {
            TransactionType aCopy = (TransactionType)this.MemberwiseClone();
            //shallow copy should be sufficient enough to generate a clone of this object...
            return aCopy;
        }

        #endregion
    }
}

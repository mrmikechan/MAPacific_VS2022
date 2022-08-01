using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace MAPacificReportUtility
{
    class GlobalSummaryBranch : INotifyPropertyChanged
    {
        public GlobalSummaryBranch() : base ()
        {
            _fundsPoolStatusList = new BindingList<FundsPoolStatusType>();
            _transactionlist = new BindingList<TransactionType>();
        }

        #region property
        private BindingList<TransactionType> _transactionlist;

        public BindingList<TransactionType> TransactionList
        {
            get { return _transactionlist; }
            set
            {
                if (value != null)
                {
                    _transactionlist = value;
                    OnPropertyChanged("TransactionList");
                }
            }
        }
        
        private BindingList<FundsPoolStatusType> _fundsPoolStatusList;
        public BindingList<FundsPoolStatusType> FundsPoolStatusList
        {
            get { return _fundsPoolStatusList; }
            set
            {
                if (value != null)
                {
                    _fundsPoolStatusList = value;
                    OnPropertyChanged("FundsPoolStatusList");
                }
            }
        }

        private string _clientid = "";
        /// <summary>
        /// string variable ClietID used to store the cliend id tag for this Branch.
        /// </summary>
        public string ClientID
        {
            get { return _clientid; }
            set
            {
                if (!_clientid.Equals(value))
                {
                    _clientid = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("ClientID");
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
    }
}

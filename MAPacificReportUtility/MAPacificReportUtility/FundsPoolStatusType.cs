using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace MAPacificReportUtility
{
    class FundsPoolStatusType : INotifyPropertyChanged
    {
        private String _fundsPoolStatus;
        public String FundsPoolStatus
        {
            get { return _fundsPoolStatus; }
            set
            {
                _fundsPoolStatus = string.IsNullOrEmpty(value) ? "" : value.Trim();
                OnPropertyChanged("FUNDS POOL STATUS");
            }
        }

        private decimal _fundsPoolBalance;
        public decimal FundsPoolBalance
        {
            get { return _fundsPoolBalance; }
            set
            {
                if (_fundsPoolBalance != value)
                {
                    _fundsPoolBalance = value;
                    OnPropertyChanged("FUNDS POOL BALANCE");
                }
            }
        }

        private string _dbcr = "CR";
        public string DBCR
        {
            get { return _dbcr; }
            set
            {
                if (!_dbcr.Equals(value))
                {
                    _dbcr = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("DBCR");
                }
            }
        }

        private decimal _accountsReported;
        public decimal AccountsReported
        {
            get { return _accountsReported; }
            set
            {
                if (_accountsReported != value)
                {
                    _accountsReported = value;
                    OnPropertyChanged("ACCOUNTS REPORTED");
                }
            }
        }

        private decimal _openAccounts;
        public decimal OpenAccounts
        {
            get { return _openAccounts; }
            set
            {
                if (_openAccounts != value)
                {
                    _openAccounts = value;
                    OnPropertyChanged("OPEN ACCOUNTS");

                }
            }
        }

        private decimal _closedAccounts;
        public decimal ClosedAccounts
        {
            get { return _closedAccounts; }
            set
            {
                if(_closedAccounts != value)
                {
                    _closedAccounts = value;
                    OnPropertyChanged("CLOSED ACCOUNTS");
                }
            }
        }

        private decimal _voidedAccounts;
        public decimal VoidedAccounts
        {
            get { return _voidedAccounts; }
            set
            {
                if (_voidedAccounts != value)
                {
                    _voidedAccounts = value;
                    OnPropertyChanged("VOIDED ACCOUNTS");

                }
            }
        }

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

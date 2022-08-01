using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace MAPacificReportUtility.excel
{
    public abstract class ExportExcel
    {
        public ExportExcel()
        {

        }

        private string _excelDirectory;
        public string ExcelDirectory
        {
            get { return _excelDirectory; }
            set
            {
                if (value != null)
                {
                    _excelDirectory = Path.GetFullPath(value);
                }
            }
        }

        private string _excelFileName;
        public string ExcelFileName
        {
            get { return _excelFileName; }
            set
            {
                if (value != null)
                {
                    _excelFileName = value + ".xls";
                }
            }
        }

        public string FullPathandFileName
        {
            get { return Path.Combine(ExcelDirectory,ExcelFileName); }
        }


    }
}

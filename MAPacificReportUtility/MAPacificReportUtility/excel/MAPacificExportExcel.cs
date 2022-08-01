using System;
using System.Collections.Generic;
using System.ComponentModel;
//using System.Text;
using System.Text.RegularExpressions;
//uncomment if using Infragistics 2012 Volume 2.
using Infragistics.Documents.Excel;
//using Infragistics.Excel;
using Infragistics.Win.UltraWinGrid;
using Infragistics.Win.UltraWinGrid.ExcelExport;
using System.Drawing;
using MAPacificReportUtility.forms;

using GraphAuthentication;

namespace MAPacificReportUtility.excel
{
    class MAPacificExportExcel : ExportExcel
    {
        //VS4591 Export Client Excel files to PDF
        //If we perform this action on the gui thread, then the app looks like
        //it is stuck. So putting it on a background worker thread instead!
        BackgroundWorker bgwExportPDF;
        FloatingProgressBar progressBarPanel = null;
        MailMessageDraftUtil mailUtil;

        GraphAuthentication.GraphHelper graphHelper;

        public MAPacificExportExcel(BindingList<ClientBranch> PrepaidCardTotalSummary, BindingList<ClientBranch> GPRSummary, BindingList<ClientBranch> GiftSummary, BindingList<ClientBranch> GPRDetail, BindingList<ClientBranch> GiftDetail, BindingList<ClientBranch> inBranchList, String GlobalSummary)
        {
            _prepaidcardtotalsummary = PrepaidCardTotalSummary;
            _gprsummary = GPRSummary;
            _giftsummary = GiftSummary;
            _gprdetail = GPRDetail;
            _giftdetail = GiftDetail;
            _globalSummary = GlobalSummary;
            _toExcelList = new List<ClientBranch>();
            _mergedBranchList = new List<ClientBranch>();
            _clientBranchList = inBranchList;
            _singleBranchReportList = new List<ClientBranch>();

            //UltraGrid to contain data for PrepaidImport data
            ultraGridExcelPrepaidImport = new UltraGridExcelExporter();
            ultraGridExcelPrepaidImport.HeaderRowExporting += new HeaderRowExportingEventHandler(ultraGridExcelPrepaidImport_HeaderRowExporting);
            ultraGridExcelPrepaidImport.CellExporting += new CellExportingEventHandler(ultraGridExcelPrepaidImport_CellExporting);
            ultraGridExcelPrepaidImport.CellExported += new CellExportedEventHandler(ultraGridExcelPrepaidImport_CellExported);
            ultraGridExcelPrepaidImport.InitializeColumn += new InitializeColumnEventHandler(ultraGridExcelPrepaidImport_InitializeColumn);

            //UltraGrid to contain data for MAP Summary
            ultraGridExcelSummaryExporter = new UltraGridExcelExporter();
            ultraGridExcelSummaryExporter.HeaderRowExporting += new HeaderRowExportingEventHandler(ultraGridExcelSummaryExporter_HeaderRowExporting);
            ultraGridExcelSummaryExporter.CellExporting += new CellExportingEventHandler(ultraGridExcelSummaryExporter_CellExporting);
            ultraGridExcelSummaryExporter.InitializeColumn += new InitializeColumnEventHandler(ultraGridExcelSummaryExporter_InitializeColumn);

            //UltraGrid to contain data for Transaction_Amount
            ultraGridExcelTransactionAmountExporter = new UltraGridExcelExporter();
            ultraGridExcelTransactionAmountExporter.InitializeColumn += new InitializeColumnEventHandler(ultraGridExcelTransactionAmountExporter_InitializeColumn);
            ultraGridExcelTransactionAmountExporter.EndExport += new EndExportEventHandler(ultraGridExcelTransactionAmountExporter_EndExport);

            //UltraGrid to contain data for Transaction_Count
            ultraGridExcelTransactionCountExporter = new UltraGridExcelExporter();
            ultraGridExcelTransactionCountExporter.InitializeColumn += new InitializeColumnEventHandler(ultraGridExcelTransactionCountExporter_InitializeColumn);
            ultraGridExcelTransactionCountExporter.EndExport += new EndExportEventHandler(ultraGridExcelTransactionCountExporter_EndExport);


            //VS4731 UltraGrid for Card_Activity_Amount
            ultraGridExcelCardActivityAmountExporter = new UltraGridExcelExporter();
            ultraGridExcelCardActivityAmountExporter.InitializeColumn += new InitializeColumnEventHandler(ultraGridExcelCardActivityAmountExporter_InitializeColumn);
            ultraGridExcelCardActivityAmountExporter.EndExport += new EndExportEventHandler(ultraGridExcelCardActivityAmountExporter_EndExport);
            
            //VS4732 UltraGrid for Card_Activity_Count
            ultraGridExcelCardActivityCountExporter = new UltraGridExcelExporter();
            ultraGridExcelCardActivityCountExporter.InitializeColumn += new InitializeColumnEventHandler(ultraGridExcelCardActivityCountExporter_InitializeColumn);
            ultraGridExcelCardActivityCountExporter.EndExport += new EndExportEventHandler(ultraGridExcelCardActivityCountExporter_EndExport);
            
            //UltraGrid to contain data for CFPRecon
            ultraGridExcelCFPReconExporter = new UltraGridExcelExporter();
            ultraGridExcelCFPReconExporter.InitializeColumn += new InitializeColumnEventHandler(ultraGridExcelCFPReconExporter_InitializeColumn);

            //UltraGrid to contain data used for individual client data
            ultraGridExcelClientReportExporter = new UltraGridExcelExporter();
            ultraGridExcelClientReportExporter.HeaderRowExporting += new HeaderRowExportingEventHandler(ultraGridExcelClientReportExporter_HeaderRowExporting);
            ultraGridExcelClientReportExporter.CellExporting += new CellExportingEventHandler(ultraGridExcelClientReportExporter_CellExporting);
            ultraGridExcelClientReportExporter.CellExported += new CellExportedEventHandler(ultraGridExcelClientReportExporter_CellExported);
            ultraGridExcelClientReportExporter.InitializeColumn += new InitializeColumnEventHandler(ultraGridExcelClientReportExporter_InitializeColumn);
            ultraGridExcelClientReportExporter.EndExport += new EndExportEventHandler(ultraGridExcelClientReportExporter_EndExport);
            dummyBranch = new ClientBranch();

            //UltragridExcelExport for ProcessWireConfirmationReport
            ultraGridExcelWireConfReport = new UltraGridExcelExporter();
            ultraGridExcelWireConfReport.InitializeColumn += new InitializeColumnEventHandler(ultraGridExcelWireConfReport_InitializeColumn);
            ultraGridExcelWireConfReport.EndExport += new EndExportEventHandler(ultraGridExcelWireConfReport_EndExport);

            //VS4591 Export Client Excel to PDF
            bgwExportPDF = new BackgroundWorker();
            bgwExportPDF.WorkerSupportsCancellation = true;
            bgwExportPDF.WorkerReportsProgress = true;
            bgwExportPDF.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgwExportPDF_RunWorkerCompleted);
            bgwExportPDF.ProgressChanged += new ProgressChangedEventHandler(bgwExportPDF_ProgressChanged);
            bgwExportPDF.DoWork += new DoWorkEventHandler(bgwExportPDF_DoWork);

            //VS4596 Create Draft Emails for user on background thread
            //7_21_2022_ Convert Background Worker Thread over to async Task. BWThread has issues with Graph Api that uses Async and Await that
            //causes the thread to end prematurely.

            //bgwSaveDraftEmail = new BackgroundWorker();
            //bgwSaveDraftEmail.WorkerSupportsCancellation = false;
            //bgwSaveDraftEmail.WorkerReportsProgress = true;
            //bgwSaveDraftEmail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgwSaveDraftEmail_RunWorkerCompleted);
            //bgwSaveDraftEmail.ProgressChanged += new ProgressChangedEventHandler(bgwSaveDraftEmail_ProgressChanged);
            //bgwSaveDraftEmail.DoWork += new DoWorkEventHandler(bgwSaveDraftEmail_DoWork);

            mailUtil = new MailMessageDraftUtil();
            graphHelper = new GraphHelper();
        }

        public void SetBindingLists(BindingList<ClientBranch> PrepaidCardTotalSummary, BindingList<ClientBranch> GPRSummary, BindingList<ClientBranch> GiftSummary, BindingList<ClientBranch> GPRDetail, BindingList<ClientBranch> GiftDetail, BindingList<ClientBranch> inBranchList, String GlobalSummary)
        {
            _prepaidcardtotalsummary = PrepaidCardTotalSummary;
            _gprsummary = GPRSummary;
            _giftsummary = GiftSummary;
            _gprdetail = GPRDetail;
            _giftdetail = GiftDetail;
            _globalSummary = GlobalSummary;
            _toExcelList = new List<ClientBranch>();
            _mergedBranchList = new List<ClientBranch>();
            _clientBranchList = inBranchList;
            _singleBranchReportList = new List<ClientBranch>();
        }
        /// <summary>
        /// Constructor for use by ProcessWireConfirmationReport class.  Essentially we can still use the
        /// excel export if we did not process the other reports.
        /// </summary>
        public MAPacificExportExcel()
        {
            //UltragridExcelExport for ProcessWireConfirmationReport
            ultraGridExcelWireConfReport = new UltraGridExcelExporter();
            ultraGridExcelWireConfReport.InitializeColumn += new InitializeColumnEventHandler(ultraGridExcelWireConfReport_InitializeColumn);
            ultraGridExcelWireConfReport.EndExport += new EndExportEventHandler(ultraGridExcelWireConfReport_EndExport);
            mailUtil = new MailMessageDraftUtil();
        }

        #region Wire Confirmation Report
        void ultraGridExcelWireConfReport_InitializeColumn(object sender, InitializeColumnEventArgs e)
        {

            if (e.Column.Key.Equals("PRC_NO"))
            {
                e.ExcelFormatStr = "[Blue]@";
            }

            if (e.Column.Key.Equals("REF"))
            {
                e.ExcelFormatStr = "[Blue]@";
            }

//            if (e.Column.Key.Equals("CREDIT_UNION"))
//            {
//                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
//            }

            if (e.Column.Key.Equals("VISA_DPS_FINANCIAL_TRANS"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,###,###,##0.00_);[Red](\"$\"#,####,###,##0.00);\"$\"#,###,###,##0.00_)";
            }

            if (e.Column.Key.Equals("EXCEPTION_TRANS"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("MISCELLANEOUS_FEES"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("VISA_ATM_REIMB_FEES"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("VISA_ATM_ISA_FEES"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("INTERLINK_REIMB_FEES"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("INTERLINK_ISA_FEES"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("GRAND_TOTAL"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,###,###,##0.00_);[Red](\"$\"#,####,###,##0.00);\"$\"#,###,###,##0.00_)";
            }

            if (e.Column.Key.Equals("WIRE_CONFIRMATION_TOTAL"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,###,###,##0.00_);[Red](\"$\"#,####,###,##0.00);\"$\"#,###,###,##0.00_)";
            }  
        }

        /// <summary>
        /// After exporting to excel, need to set the column widths and wrap some header text.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelWireConfReport_EndExport(object sender, EndExportEventArgs e)
        {
            //wrap the header text for the grouping.
            wsWireConfReport.Rows[0].Cells[3].CellFormat.WrapText = ExcelDefaultableBoolean.True;
            wsWireConfReport.Rows[0].Cells[4].CellFormat.WrapText = ExcelDefaultableBoolean.True;
            wsWireConfReport.Rows[0].Cells[5].CellFormat.WrapText = ExcelDefaultableBoolean.True;
            wsWireConfReport.Rows[0].Cells[6].CellFormat.WrapText = ExcelDefaultableBoolean.True;
            wsWireConfReport.Rows[0].Cells[7].CellFormat.WrapText = ExcelDefaultableBoolean.True;
            wsWireConfReport.Rows[0].Cells[8].CellFormat.WrapText = ExcelDefaultableBoolean.True;
            wsWireConfReport.Rows[0].Cells[9].CellFormat.WrapText = ExcelDefaultableBoolean.True;
            wsWireConfReport.Rows[0].Cells[10].CellFormat.WrapText = ExcelDefaultableBoolean.True;
            wsWireConfReport.Rows[0].Cells[11].CellFormat.WrapText = ExcelDefaultableBoolean.True;
            
            //set the excel column height for the group header.
            //the height (not pixels) numbers are retrieved from the excel worksheet and multiply it by 20 to get
            //the height and width values
            wsWireConfReport.Rows[0].Height = 62 * 20;

            ////set the excel column width times 256
            wsWireConfReport.Columns[0].Width = 9 * 256;
            wsWireConfReport.Columns[1].Width = 11 * 256;
            wsWireConfReport.Columns[2].Width = 44 * 256;
            wsWireConfReport.Columns[3].Width = 18 * 256;
            wsWireConfReport.Columns[4].Width = 17 * 256;
            wsWireConfReport.Columns[5].Width = 17 * 256;
            wsWireConfReport.Columns[6].Width = 17 * 256;
            wsWireConfReport.Columns[7].Width = 17 * 256;
            wsWireConfReport.Columns[8].Width = 17 * 256;
            wsWireConfReport.Columns[9].Width = 17 * 256;
            wsWireConfReport.Columns[10].Width = 15 * 256;
            wsWireConfReport.Columns[11].Width = 15 * 256;
            
            //double underline the last row
            wsWireConfReport.Rows[e.CurrentRowIndex].CellFormat.Font.UnderlineStyle = FontUnderlineStyle.DoubleAccounting;
            
        }

        #endregion

        #region prepaidimport excel exporter

        /// <summary>
        /// Custom formatting of Header row to perform the following:
        /// Delete extra row between Client ID and “Transaction” row.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelPrepaidImport_HeaderRowExporting(object sender, HeaderRowExportingEventArgs e)
        {
            if (e.GridRow.Band.Index != 0)
            {
                if (e.GridRow.HasPrevSibling() == false)
                {
                    e.CurrentRowIndex -= 1;
                }
            }
        }

        /// <summary>
        /// MAP rational for flipping values:
        /// We need to change CR for DR because whenever you see CR in those reports, it is actually debit for credit Unions. 
        /// CR means that funds were credited to prepaid card, thus we have to go out and take money. So, we perform debit to Credit Union’s account. 
        /// For credit union it will be much understandable, if thy will see opposite to what is actually in the report.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelPrepaidImport_CellExporting(object sender, CellExportingEventArgs e)
        {
            //logic to flip the values for CR to DR and if DR flip to CR per reuqest by MAP.
            if (e.Value.ToString().Equals("CR", StringComparison.CurrentCultureIgnoreCase))
            {
                e.Value = "DR";
                return;
            }

            if (e.Value.ToString().Equals("DR", StringComparison.CurrentCultureIgnoreCase))
            {
                e.Value = "CR";
                return;
            }
        }

        /// <summary>
        /// MAP wants to add some custom formatting to the Excel sheet:
        /// Make line “Total Load/Unload activity bold with text size 11 for SubClient ID XXXX000. 
        /// Center columns C, D, and E (TransactionCount,TransactionAmount, DBCR1)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelPrepaidImport_CellExported(object sender, CellExportedEventArgs e)
        {
            //uncomment if using infragistics 2012 vol. 2
            Infragistics.Documents.Excel.IWorksheetCellFormat cfCellFmt;
            //Infragistics.Excel.IWorksheetCellFormat cfCellFmt;

            //logic to handle the bolding of the Total Load/Unload row of data.
            try
            {
                value = e.GridRow.Cells["Transaction"].Value.ToString();
            }
            catch (Exception ex)
            {
                value = null;
            }
            if (value != null && value.Equals("*TOTAL LOAD/UNLOAD ACTIVITY", StringComparison.CurrentCultureIgnoreCase))
            {
                try
                {
                    // Set format property for the font to bold
                    dummyBranch.ClientID = e.GridRow.ParentRow.Cells["ClientID"].Value.ToString();
                    dummyBranch.RelationalParentSummary = e.GridRow.ParentRow.Cells["RelationalParentSummary"].Value.ToString();

                    //we don't want to add bold font to summary branch that has been merged with another clientID.
                    if (_prepaidcardtotalsummary.Contains(dummyBranch) && dummyBranch.RelationalParentSummary.Length == 0)
                    {
                        int iRdex = e.CurrentRowIndex;
                        int iCdex = e.CurrentColumnIndex;
                        cfCellFmt = e.CurrentWorksheet.Rows[iRdex].Cells[iCdex].CellFormat;
                        cfCellFmt.Font.Bold = ExcelDefaultableBoolean.True; //Make bold
                        cfCellFmt.Font.Height = 215;                        //Make font size 11

                        // Apply the formatting, this step minimizes the number of
                        // Worksheet Font objects that need to be instantiated.
                        e.CurrentWorksheet.Rows[iRdex].Cells[iCdex].CellFormat.SetFormatting(cfCellFmt);
                    }
                }//reach exception in handling ultragrid for globaldatasummary table because in this band there is no ClientID or RelationalParentSummary so just gobble up the error
                catch (Exception ex2)
                {
                }
            }
        }

        /// <summary>
        /// Initialize the columns in the excel report so that they have the correct formatting.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelPrepaidImport_InitializeColumn(object sender, InitializeColumnEventArgs e)
        {
            if (e.Column.Key.Equals("TransactionCount", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "#,##0";
            }

            if (e.Column.Key.Equals("TransactionAmount", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "$#,##0.00";
            }

            if (e.Column.Key.Equals("FundsPoolBalance", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "$#,##0.00";
            }

            System.Console.WriteLine(e.ToString());
        }

        /// <summary>
        /// At the end of Excel Export process, customer would like to have the contents in the worksheet to be expanded instead of collapsed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelPrepaidImport_EndExport(object sender, EndExportEventArgs e)
        {
            foreach (WorksheetRow xlsRow in e.CurrentWorksheet.Rows)
            {
                xlsRow.Hidden = false;
            }
        }

        #endregion

        #region excel exporter events for formatting excel columns in CFPRecon

        /// <summary>
        /// MAPacific wants blue font for positive values, red font for negative values and black font for zero values.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelCFPReconExporter_InitializeColumn(object sender, InitializeColumnEventArgs e)
        {
            
            if (e.Column.Key.Equals("Debit1"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";

            }

            if (e.Column.Key.Equals("Debit2"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";

            }

            if (e.Column.Key.Equals("Credit1"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";

            }

            if (e.Column.Key.Equals("Credit2"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";

            }
        }

        #endregion

        #region excel exporter events for formatting excel columns in Transaction Amount and Transaction Count reports.

        /// <summary>
        /// MAPacific want blue font color for positive value and balck font color for values that are zero.  There are no negative values
        /// in this report.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelTransactionCountExporter_InitializeColumn(object sender, InitializeColumnEventArgs e)
        {
            if (e.Column.Key.Equals("LOADS_FI_FUNDS_TRNSFER"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("RELOAD_FI_FUNDS_TRNSFER"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("LOADS_MERCH_POS_FUNDING"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            //VS4637 Add Loads Bypass
            if (e.Column.Key.Equals("LOADS_BYPASS"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("RELOADS_BYPASS"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            //VS4637 Add Loads Merchant Acq
            if (e.Column.Key.Equals("LOADS_MERCHANT_ACQ"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }            
            
            if (e.Column.Key.Equals("RELOADS_MERCHANT_ACQ"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            //VS4637 Add Unloads Merchant Acq
            if (e.Column.Key.Equals("UNLOADS_MERCHANT_ACQ"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("UNLOADS_FI_FUNDS_TRNSFR"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("UNLOADS_BYPASS"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("MANUAL_ADJUSTMENTS"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("TOTAL_TRANSACTION_COUNT"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }
        }

        /// <summary>
        /// ultraGridExcelTransactionAmountExporter_InitializeColumn event handler is used to format the Excel columns.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelTransactionAmountExporter_InitializeColumn(object sender, InitializeColumnEventArgs e)
        {
            if (e.Column.Key.Equals("LOADS_FI_FUNDS_TRNSFER"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            //VS4730
            if (e.Column.Key.Equals("AVERAGE_LOAD"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("RELOAD_FI_FUNDS_TRNSFER"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("LOADS_MERCH_POS_FUNDING"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            //VS4637 Need to format the additional column added for Loads Bypass
            if (e.Column.Key.Equals("LOADS_BYPASS"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("RELOADS_BYPASS"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            //VS4637 Need to format the additional column added for Loads Merchant Acq
            if (e.Column.Key.Equals("LOADS_MERCHANT_ACQ"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("RELOADS_MERCHANT_ACQ"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            //VS4637 Need to format the additional column added for Unloads Merchant Acq
            if (e.Column.Key.Equals("UNLOADS_MERCHANT_ACQ"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("UNLOADS_FI_FUNDS_TRNSFR"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("UNLOADS_BYPASS"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("MANUAL_ADJUSTMENTS"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("TOTAL_TRANSACTION_AMOUNT"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            //VS6159 Add new columns ACH_DIRECT_DEPOSIT and RELOADS_MONEY_TSFR_RCVD
            if (e.Column.Key.Equals("ACH_DIRECT_DEPOSIT"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("RELOADS_MONEY_TSFR_RCVD"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }
        }

        /// <summary>
        /// EndExport event allows us to reset the Excel Format String values for cells in the Total row at the
        /// bottom of the report.  We need to reset these format values so that the font color for the totals stay
        /// white in color and not blue.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelTransactionAmountExporter_EndExport(object sender, EndExportEventArgs e)
        {
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[2].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[3].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[4].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[5].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[6].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[7].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[8].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[9].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[10].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            //VS4637 Need to compensate for the extra two columns added
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[11].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[12].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[13].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
        }

        /// <summary>
        /// EndExport event allows us to reset the Excel Format String values for cells in the Total row at the
        /// bottom of the report.  We need to reset these format values so that the font color for the totals stay
        /// white in color and not blue.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelTransactionCountExporter_EndExport(object sender, EndExportEventArgs e)
        {
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[2].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[3].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[4].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[5].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[6].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[7].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[8].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[9].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[10].CellFormat.FormatString = "#,##0";
            //VS4637 Need to compensate for the extra two columns added
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[11].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[12].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[13].CellFormat.FormatString = "#,##0";
            
        }

        #endregion

        #region excel export events for formatting excel columns in Card_Activity_Amount and Card_Activity_Count
        //VS4731 and VS4732 events for Card Activity
        void ultraGridExcelCardActivityCountExporter_EndExport(object sender, EndExportEventArgs e)
        {
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[2].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[3].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[4].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[5].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[6].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[7].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[8].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[9].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[10].CellFormat.FormatString = "#,##0";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[11].CellFormat.FormatString = "#,##0";
        }

        void ultraGridExcelCardActivityAmountExporter_EndExport(object sender, EndExportEventArgs e)
        {
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[2].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[3].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[4].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[5].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[6].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[7].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[8].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[9].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[10].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            e.CurrentWorksheet.Rows[e.CurrentRowIndex - 1].Cells[11].CellFormat.FormatString = "[White][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
        }

        void ultraGridExcelCardActivityCountExporter_InitializeColumn(object sender, InitializeColumnEventArgs e)
        {
            if (e.Column.Key.Equals("LOAD_DISPUTES"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("PURCHASES_QUASI_CASH"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("PURCHASES_WITH_CASH_BACK"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("AFT_AA/PP"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("PURCHASE_RETURNS"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("MANUAL_CASH"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("ATM_CASH"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("BALANCE_INQUIRIES"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("REWARDS"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("TOTAL_CARD_ACTIVITY"))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }
        }

        void ultraGridExcelCardActivityAmountExporter_InitializeColumn(object sender, InitializeColumnEventArgs e)
        {
            if (e.Column.Key.Equals("LOAD_DISPUTES"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("PURCHASES_QUASI_CASH"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("PURCHASES_WITH_CASH_BACK"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("AFT_AA/PP"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("PURCHASE_RETURNS"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("MANUAL_CASH"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("ATM_CASH"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("BALANCE_INQUIRIES"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("REWARDS"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("TOTAL_CARD_ACTIVITY"))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }
        }
        #endregion

        #region excel exporter events for creating individual excel reports


        /// <summary>
        /// Initialize the columns in the excel report so that they have the correct formatting.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelClientReportExporter_InitializeColumn(object sender, InitializeColumnEventArgs e)
        {
            if (e.Column.Key.Equals("TransactionCount", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "#";
            }

            if (e.Column.Key.Equals("TransactionAmount", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "$#,##0.00";
            }

            if (e.Column.Key.Equals("FundsPoolBalance", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "$#,##0.00";
            }

            System.Console.WriteLine(e.ToString());
        }

        /// <summary>
        /// MAP wants to add some custom formatting to the Excel sheet:
        /// Make line “Total Load/Unload activity bold with text size 11 for SubClient ID XXXX000. 
        /// Center columns C, D, and E (TransactionCount,TransactionAmount, DBCR1)
        /// bold text ClientID Names only for client branches that are grouped in the XXXX000, XXXX100, ZXXX100 categories. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelClientReportExporter_CellExported(object sender, CellExportedEventArgs e)
        {
            //uncomment if using infragistics 2012 vol. 2
            Infragistics.Documents.Excel.IWorksheetCellFormat cfCellFmt;
            //Infragistics.Excel.IWorksheetCellFormat cfCellFmt;

            //add extra logic to bold the Client ID Name for XXXX000, XXX100, and ZXXX100 types.
            try
            {
                value = e.GridRow.Cells["ClientID"].Value.ToString();
            }
            catch (Exception ex)
            {
                value = null;
            }

            if (value != null && value.Length > 0)
            {
                try
                {
                    dummyBranch.ClientID = value;
                    dummyBranch.RelationalParentSummary = e.GridRow.Cells["RelationalParentSummary"].Value.ToString();

                    if (_prepaidcardtotalsummary.Contains(dummyBranch) || _gprsummary.Contains(dummyBranch) || _giftsummary.Contains(dummyBranch))
                    {
                        if (dummyBranch.RelationalParentSummary.Length == 0)
                        {
                            //find the row index to apply the formatting. The column will always start at the 0 index...
                            int iRdex = e.CurrentRowIndex;
                            int iCdex = e.CurrentColumnIndex;

                            cfCellFmt = e.CurrentWorksheet.Rows[iRdex].Cells[iCdex].CellFormat;
                            cfCellFmt.Font.Bold = ExcelDefaultableBoolean.True; //Make bold
                            cfCellFmt.Font.Height = 215;                        //Make font size 11

                            // Apply the formatting, this step minimizes the number of
                            // Worksheet Font objects that need to be instantiated.
                            e.CurrentWorksheet.Rows[iRdex].Cells[iCdex].CellFormat.SetFormatting(cfCellFmt);
                        }
                    }
                }
                catch (Exception ex3)
                { }
            }

            //logic to handle the bolding of the Total Load/Unload row of data.
            try
            {
                value = e.GridRow.Cells["Transaction"].Value.ToString();
            }
            catch (Exception ex)
            {
                value = null;
            }
            if (value != null && value.Equals("*TOTAL LOAD/UNLOAD ACTIVITY", StringComparison.CurrentCultureIgnoreCase))
            {
                try
                {
                    // Set format property for the font to bold
                    dummyBranch.ClientID = e.GridRow.ParentRow.Cells["ClientID"].Value.ToString();
                    dummyBranch.RelationalParentSummary = e.GridRow.ParentRow.Cells["RelationalParentSummary"].Value.ToString();

                    //we don't want to add bold font to summary branch that has been merged with another clientID.
                    if (_prepaidcardtotalsummary.Contains(dummyBranch) || _gprsummary.Contains(dummyBranch) || _giftsummary.Contains(dummyBranch))
                    {
                        //we don't want to add bold font to summary branch that has been merged with another clientID.
                        if (dummyBranch.RelationalParentSummary.Length == 0)
                        {
                            int iRdex = e.CurrentRowIndex;
                            int iCdex = e.CurrentColumnIndex;
                            cfCellFmt = e.CurrentWorksheet.Rows[iRdex].Cells[iCdex].CellFormat;
                            cfCellFmt.Font.Bold = ExcelDefaultableBoolean.True; //Make bold
                            cfCellFmt.Font.Height = 215;                        //Make font size 11

                            // Apply the formatting, this step minimizes the number of
                            // Worksheet Font objects that need to be instantiated.
                            e.CurrentWorksheet.Rows[iRdex].Cells[iCdex].CellFormat.SetFormatting(cfCellFmt);
                        }
                    }
                }//reach exception in handling ultragrid for globaldatasummary table because in this band there is no ClientID or RelationalParentSummary so just gobble up the error
                catch (Exception ex2)
                {
                }
            }
        }

        /// <summary>
        /// MAP rational for flipping values:
        /// We need to change CR for DR because whenever you see CR in those reports, it is actually debit for credit Unions. 
        /// CR means that funds were credited to prepaid card, thus we have to go out and take money. So, we perform debit to Credit Union’s account. 
        /// For credit union it will be much understandable, if thy will see opposite to what is actually in the report.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelClientReportExporter_CellExporting(object sender, CellExportingEventArgs e)
        {
            //logic to flip the values for CR to DR and if DR flip to CR per reuqest by MAP.
            if (e.Value.ToString().Equals("CR", StringComparison.CurrentCultureIgnoreCase))
            {
                e.Value = "DR";
                return;
            }

            if (e.Value.ToString().Equals("DR", StringComparison.CurrentCultureIgnoreCase))
            {
                e.Value = "CR";
                return;
            }
        }

        /// <summary>
        /// Custom formatting of Header row to perform the following:
        /// Delete extra row between Client ID and “Transaction” row.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelClientReportExporter_HeaderRowExporting(object sender, HeaderRowExportingEventArgs e)
        {
            if (e.GridRow.Band.Index != 0)
            {
                if (e.GridRow.HasPrevSibling() == false)
                {
                    e.CurrentRowIndex -= 1;
                }
            }
        }

        //VS3599N Expand the rows(bands) in the excel report by default...
        /// <summary>
        /// At the end of the ultragrid export to excel, have all the rows (bands) to be expanded initially
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelClientReportExporter_EndExport(object sender, EndExportEventArgs e)
        {
            foreach (WorksheetRow row in e.CurrentWorksheet.Rows)
            {
                row.Hidden = false;
            }
        }   
        #endregion

        #region excel export events for creating Map Summary
        /// <summary>
        /// MAP rational for flipping values:
        /// We need to change CR for DR because whenever you see CR in those reports, it is actually debit for credit Unions. 
        /// CR means that funds were credited to prepaid card, thus we have to go out and take money. So, we perform debit to Credit Union’s account. 
        /// For credit union it will be much understandable, if thy will see opposite to what is actually in the report.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelSummaryExporter_CellExporting(object sender, CellExportingEventArgs e)
        {
            //logic to flip the values for CR to DR and if DR flip to CR per reuqest by MAP.
            if (e.Value.ToString().Equals("CR", StringComparison.CurrentCultureIgnoreCase))
            {
                e.Value = "DR";
                return;
            }

            if (e.Value.ToString().Equals("DR", StringComparison.CurrentCultureIgnoreCase))
            {
                e.Value = "CR";
                return;
            }
        }


        /// <summary>
        /// Initialize the columns in the excel report so that they have the correct formatting.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelSummaryExporter_InitializeColumn(object sender, InitializeColumnEventArgs e)
        {
            if (e.Column.Key.Equals("TransactionCount", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("TransactionAmount", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("FeeAmount", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("TotalAmount", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("FundsPoolBalance", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "[Blue][>0]\"$\"#,##0.00_);[Red](\"$\"#,##0.00);\"$\"#,##0.00_)";
            }

            if (e.Column.Key.Equals("AccountsReported", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("OpenAccounts", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("ClosedAccounts", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            if (e.Column.Key.Equals("VoidedAccounts", StringComparison.CurrentCultureIgnoreCase))
            {
                e.ExcelFormatStr = "[Blue][>0]#,##0;#,##0";
            }

            //System.Console.WriteLine(e.ToString());
        }

        /// <summary>
        /// Custom formatting of Header row to perform the following:
        /// Delete extra row between Client ID and “Transaction” row.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelSummaryExporter_HeaderRowExporting(object sender, HeaderRowExportingEventArgs e)
        {
            if (e.GridRow.Band.Index != 0)
            {
                if (e.GridRow.HasPrevSibling() == false)
                {
                    e.CurrentRowIndex -= 1;
                }
            }
        }

        /// <summary>
        /// At the end of Excel Export process, customer would like to have the contents in the worksheet to be expanded instead of collapsed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultraGridExcelSummaryExporter_EndExport(object sender, EndExportEventArgs e)
        {
            foreach (WorksheetRow xlsRow in e.CurrentWorksheet.Rows)
            {
                xlsRow.Hidden = false;
            }
        }
        #endregion

        #region member variables/properties

        BindingList<ClientBranch> _prepaidcardtotalsummary;
        BindingList<ClientBranch> _gprsummary;
        BindingList<ClientBranch> _giftsummary;
        BindingList<ClientBranch> _gprdetail;
        BindingList<ClientBranch> _giftdetail;
        BindingList<ClientBranch> _clientBranchList;

        String _globalSummary = "";
        //Excel Exporter for the whole report parsed from Visa DPS
        UltraGridExcelExporter ultraGridExcelSummaryExporter;
        UltraGridExcelExporter ultraGridExcelPrepaidImport;
        UltraGridExcelExporter ultraGridExcelTransactionAmountExporter;
        UltraGridExcelExporter ultraGridExcelTransactionCountExporter;
        UltraGridExcelExporter ultraGridExcelCFPReconExporter;
        //VS4731 Add new worksheet CARD_ACTIVITY_AMOUNT
        UltraGridExcelExporter ultraGridExcelCardActivityAmountExporter;
        //VS4732 Add new worksheet CARD_ACTIVITY_COUNT
        UltraGridExcelExporter ultraGridExcelCardActivityCountExporter;

        //Excel Exporter for the individual client branch report parsed from the
        //whole report
        UltraGridExcelExporter ultraGridExcelClientReportExporter;
        //Excel objects
        Workbook wb;
        Worksheet ws;
        Worksheet ws2;
        Worksheet ws3;
        Worksheet ws4;
        Worksheet ws5;
        //VS4731 and VS4732
        Worksheet ws6;
        Worksheet ws7;
        Worksheet wsWireConfReport;

        UltraGridExcelExporter ultraGridExcelWireConfReport;
        
        string value = ""; //string variable used for manipulating string objects or holding temp data...
        ClientBranch dummyBranch; //dummyBranch used for processing data for ClientBranch objects

        //List of dataobjects in a specific order that is used to export the data out to excel.
        private List<ClientBranch> _toExcelList;
        public List<ClientBranch> ToExcelList
        {
            get { return _toExcelList; }
        }

        //List used for preparing the individual client branch reports
        private List<ClientBranch> _singleBranchReportList;

        //List of ClientBranch objects that have been gobbled up by another branch.
        List<ClientBranch> _mergedBranchList;

        #endregion

        public void SortDataIntoMAPFormat()
        {
            //check to make sure we have a complete parsed report to work on by
            //knowing that if we have the string content for the Global Summary then we
            //have successfully completed the parsing of the report.
            if (_globalSummary.Length > 0)
            {
                Regex pattern;
                string parentBranchName = "";
                
                //Note phase 2 development change: move Gift Summary down to after GPR Details...
                //format of excel report:
                //XXX000  - Prepaid Card Total Summary for a parent branch
                //XXX100  - GPR Summary for client ID falling under parent branch  XXX000  Follows after PRe Paid Total 
                //XXXX101 - GPR Detail for client ID falling under parent branch. Follows after GPR Summary
                //ZXXX100 - Gift Summary for client ID falling under parent branch XXX000. Follows after GPR Summary
                //ZXXX101 - Gift Detail for client ID falling under parent branch. Follows after GPR Detail
                //No more branches then Global Summary data.

                //go through the branchlist and find which client branch has a relationship with another branch...
                foreach (ClientBranch mBranch in _clientBranchList)
                {
                    if (mBranch.RelationalParentSummary.Length > 0)
                    {
                        _mergedBranchList.Add(mBranch);
                    }
                }

                try
                {
                    //start assembling the list starting with Prepaid Card Total Summary
                    foreach (ClientBranch mPrepaidTotalBranch in _prepaidcardtotalsummary)
                    {
                        //we have a valid ClientBranch from Prepaid Card Total Summary list
                        if (mPrepaidTotalBranch.TransactionList.Count > 0)
                        {
                            //we want only branches that do not have any relationalsummary data
                            if(mPrepaidTotalBranch.RelationalParentSummary.Length == 0)
                            {
                                _toExcelList.Add(mPrepaidTotalBranch);

                                //store the partial branch ID into variable for use in comparing sub branch names...
                                //first check if the name is all alphabet or is alphanumeric
                                pattern = new Regex("[A-Za-z]+");
                                if (pattern.IsMatch(mPrepaidTotalBranch.ClientID))
                                {
                                    //we have a name that is all alpha characters and we know that
                                    //sub client branch names consist of 3 to 4 chars and then numbers
                                    if (mPrepaidTotalBranch.ClientID.Length > 3)
                                    {
                                        parentBranchName = mPrepaidTotalBranch.ClientID.Substring(0, 4);
                                    }
                                    else
                                    {
                                        parentBranchName = mPrepaidTotalBranch.ClientID.Substring(0, 3);
                                    }
                                }
                                else
                                {
                                    //we have alphanumeric client id name so we need to parse the letters away from the last 3 digits
                                    parentBranchName = mPrepaidTotalBranch.ClientID.Substring(0, mPrepaidTotalBranch.ClientID.Length - 3);
                                }

                                //start searching for the children branches.
                                //look in XXX100 GPR summary list
                                foreach (ClientBranch mGPRSummaryBranch in _gprsummary)
                                {
                                    //we don't want any client branch with relationalsummary data yet.
                                    if (mGPRSummaryBranch.RelationalParentSummary.Length == 0)
                                    {
                                        //we want client branches that start with the same name...
                                        //so first we need to check the parent branch name
                                        if(mGPRSummaryBranch.ClientID.StartsWith(parentBranchName, StringComparison.CurrentCultureIgnoreCase))
                                        {
                                            //we have a match and there is data in there lets add it the excel list..
                                            if(mGPRSummaryBranch.TransactionList.Count > 0)
                                                _toExcelList.Add(mGPRSummaryBranch);
                                        }
                                    }
                                } 

                                //look in XXXX101 GPR detail list
                                foreach (ClientBranch mGPRDetailBranch in _gprdetail)
                                {
                                    //we don't want any client branch with relationalsummary data yet.
                                    if (mGPRDetailBranch.RelationalParentSummary.Length == 0)
                                    {
                                        //we want client branches that start with the same name...
                                        //so first we need to check the parent branch name
                                        if (mGPRDetailBranch.ClientID.StartsWith(parentBranchName, StringComparison.CurrentCultureIgnoreCase))
                                        {
                                            //we have a match and there is data in there lets add it the excel list..
                                            if (mGPRDetailBranch.TransactionList.Count > 0)
                                                _toExcelList.Add(mGPRDetailBranch);
                                        }
                                    }
                                }
                                //VS4142 Remove Z from Gift names because of the changes implemented by VISA.
                                //ZXXX101 --> XXX100
                                //look in ZXXX100 GIFT summary list
                                foreach (ClientBranch mGiftSummaryBranch in _giftsummary)
                                {
                                    //we don't want any client branch with relationalsummary data yet.
                                    if (mGiftSummaryBranch.RelationalParentSummary.Length == 0)
                                    {
                                        //we want client branches that start with the same name...
                                        //so first we need to check the parent branch name
                                        if (mGiftSummaryBranch.ClientID.StartsWith(parentBranchName, StringComparison.CurrentCultureIgnoreCase))
                                        {
                                            //we have a match and there is data in there lets add it the excel list..
                                            if (mGiftSummaryBranch.TransactionList.Count > 0)
                                                _toExcelList.Add(mGiftSummaryBranch);
                                        }
                                    }
                                }
                                //VS4142 Remove Z from Gift names because of the changes implemented by VISA.
                                //ZXXX101 --> XXX101
                                //look in ZXXX101 GIFT detail list
                                foreach (ClientBranch mGiftDetailBranch in _giftdetail)
                                {
                                    //we don't want any client branch with relationalsummary data yet.
                                    if (mGiftDetailBranch.RelationalParentSummary.Length == 0)
                                    {
                                        //we want client branches that start with the same name...
                                        //so first we need to check the parent branch name
                                        if (mGiftDetailBranch.ClientID.StartsWith(parentBranchName, StringComparison.CurrentCultureIgnoreCase))
                                        {
                                            //we have a match and there is data in there lets add it the excel list..
                                            if (mGiftDetailBranch.TransactionList.Count > 0)
                                                _toExcelList.Add(mGiftDetailBranch);
                                        }
                                    }
                                }

                                //now check the mergecbranch list to see if we need to add the Client Branches with relationship data into here
                                foreach (ClientBranch mMergedBranch in _mergedBranchList)
                                {
                                    if(mMergedBranch.RelationalParentSummary.StartsWith(parentBranchName, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        if(mMergedBranch.TransactionList.Count > 0)
                                        {
                                            _toExcelList.Add(mMergedBranch);
                                        }
                                    }
                                }

                            }
                        }
                    } 
                }catch(Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Export to Excel Error", System.Windows.Forms.MessageBoxButtons.OK);
                }//finished sorting the list

            }
        }

        /// <summary>
        /// Export the contents within the UltraGrid into an Excel file. The contents of the UltraGrid contains the report
        /// parsed from processing the VisaDPS report.
        /// </summary>
        /// <param name="clientBranchIDGrid">The ultragrid containing all ClientBranchID with transaction information.</param>
        /// <param name="globalTotalSummaryGrid">The ultragrid containing the Global Total summary parsed from the very end of the VisaDPS report.</param>
        /// <param name="transactionAmountGrid">The ultragrid containing the Transaction_Amount worksheet information.</param>
        /// <param name="transactionCountGrid">The ultragrid containing the Transaction_Count worksheet information.</param>
        /// <param name="cfpReconGrid">The ultragrid containing the CFPRecon worsheet information.</param>
        /// <param name="fileName"></param>
        /// <param name="cardActivityAmount">ultragrid containing the Card Activity Amount worksheet information</param> 
        /// <param name="cardActivityCount">ultragrid containing the card activity count worksheet information</param>
        public void ExportToExcel(UltraGrid clientBranchIDGrid, UltraGrid globalTotalSummaryGrid, UltraGrid transactionAmountGrid, UltraGrid transactionCountGrid, UltraGrid cfpReconGrid, string fileName, UltraGrid cardActivityAmountGrid, UltraGrid cardActivityCountGrid)
        {   
            ExtractFileName(fileName);
            wb = new Workbook();
            //VS4142 Need to label the worksheet for PrepaidImport based on the report type...
            if (fileName.Contains("Gift"))
            {
                ws = wb.Worksheets.Add("GiftImport");
            }
            else
            {
                ws = wb.Worksheets.Add("PrepaidImport");
            }
            
            ws2 = wb.Worksheets.Add("Transaction_Amount");
            ws3 = wb.Worksheets.Add("Transaction_Count");
            //VS4731 4732
            ws4 = wb.Worksheets.Add("Card_Activity_Amount");
            ws5 = wb.Worksheets.Add("Card_Activity_Count");

            ws6 = wb.Worksheets.Add("CFPRecon");
            ws7 = wb.Worksheets.Add("MAP Summary");
            //Add event listener to determine when the Export to Excel has completed. 
            ultraGridExcelPrepaidImport.EndExport += new EndExportEventHandler(ultraGridExcelPrepaidImport_EndExport);
            ultraGridExcelPrepaidImport.Export(clientBranchIDGrid, ws);
            ws.DisplayOptions.ShowGridlines = false;
            ws.Columns[2].Width = 18 * 256;

            ultraGridExcelTransactionAmountExporter.Export(transactionAmountGrid, ws2);
            ws2.DisplayOptions.ShowGridlines = false;
            ws2.Columns[1].CellFormat.Indent = 1;
            ws2.DisplayOptions.PanesAreFrozen = true;
            ws2.DisplayOptions.FrozenPaneSettings.FrozenRows = 1;
            ws2.DisplayOptions.FrozenPaneSettings.FrozenColumns = 2;

            ultraGridExcelTransactionCountExporter.Export(transactionCountGrid, ws3);
            ws3.DisplayOptions.ShowGridlines = false;
            ws3.Columns[1].CellFormat.Indent = 1;
            ws3.DisplayOptions.PanesAreFrozen = true;
            ws3.DisplayOptions.FrozenPaneSettings.FrozenRows = 1;
            ws3.DisplayOptions.FrozenPaneSettings.FrozenColumns = 2;

            //VS4731
            ultraGridExcelCardActivityAmountExporter.Export(cardActivityAmountGrid, ws4);
            ws4.DisplayOptions.ShowGridlines = false;
            ws4.Columns[1].CellFormat.Indent = 1;
            ws4.DisplayOptions.PanesAreFrozen = true;
            ws4.DisplayOptions.FrozenPaneSettings.FrozenRows = 1;
            ws4.DisplayOptions.FrozenPaneSettings.FrozenColumns = 2;

            //VS4732
            ultraGridExcelCardActivityCountExporter.Export(cardActivityCountGrid, ws5);
            ws5.DisplayOptions.ShowGridlines = false;
            ws5.Columns[1].CellFormat.Indent = 1;
            ws5.DisplayOptions.PanesAreFrozen = true;
            ws5.DisplayOptions.FrozenPaneSettings.FrozenRows = 1;
            ws5.DisplayOptions.FrozenPaneSettings.FrozenColumns = 2;

            ultraGridExcelCFPReconExporter.Export(cfpReconGrid, ws6);
            ws6.DisplayOptions.ShowGridlines = false;
            ws6.Columns[0].Width = 46 * 256;
            ws6.Columns[0].CellFormat.Indent = 1;
            //Add event listener to determine when the Export to Excel has completed. 
            ultraGridExcelSummaryExporter.EndExport += new EndExportEventHandler(ultraGridExcelSummaryExporter_EndExport);
            ultraGridExcelSummaryExporter.Export(globalTotalSummaryGrid, ws7);
            ws7.DisplayOptions.ShowGridlines = false;
            ws7.Columns[0].Width = 10 * 256;
            ws7.Columns[1].CellFormat.Indent = 1;
            wb.Save(FullPathandFileName);
        }

        //VS4183 Export the CFPRecon and MapSummary to excel worksheet.
        /// <summary>
        /// Export the contents within the UltraGrid into an Excel file. The contents of the UltraGrid contains the report
        /// parsed from processing the VisaDPS report.
        /// </summary>
        /// <param name="globalTotalSummaryGrid">The ultragrid containing the Global Total summary parsed from the very end of the VisaDPS report.</param>
        /// <param name="cfpReconGrid">The ultragrid containing the CFPRecon worsheet information.</param>
        /// <param name="fileName"></param>
        public void ExportToExcel(UltraGrid globalTotalSummaryGrid, UltraGrid cfpReconGrid, string fileName)
        {
            ExtractFileName(fileName);
            wb = new Workbook();

            ws6 = wb.Worksheets.Add("CFPRecon");
            ws7 = wb.Worksheets.Add("MAP Summary");
            //Add event listener to determine when the Export to Excel has completed. 
            ultraGridExcelPrepaidImport.EndExport += new EndExportEventHandler(ultraGridExcelPrepaidImport_EndExport);
            //VS4820 When adding the last new feature, forgot to update this section of code to reference the correct ws6 and ws7 worksheets for CFPRecon and Map Summary.
            ultraGridExcelCFPReconExporter.Export(cfpReconGrid, ws6);
            ws6.DisplayOptions.ShowGridlines = false;
            ws6.Columns[0].Width = 46 * 256;
            ws6.Columns[0].CellFormat.Indent = 1;
            //Add event listener to determine when the Export to Excel has completed. 
            ultraGridExcelSummaryExporter.EndExport += new EndExportEventHandler(ultraGridExcelSummaryExporter_EndExport);
            ultraGridExcelSummaryExporter.Export(globalTotalSummaryGrid, ws7);
            ws7.DisplayOptions.ShowGridlines = false;
            ws7.Columns[0].Width = 10 * 256;
            ws7.Columns[1].CellFormat.Indent = 1;
            wb.Save(FullPathandFileName);
        }

        public void ExportWireConfirmationReport(UltraGrid wcrGrid, string date)
        {
            ExcelFileName = "Wire_Confirmation_Report_" + date;
            Workbook wb1 = new Workbook();
            wsWireConfReport = wb1.Worksheets.Add("WIRES");
            wsWireConfReport.DisplayOptions.ShowGridlines = false;
            wsWireConfReport.DisplayOptions.PanesAreFrozen = true;
            //set the row number where scrolling will take place and leave the header present.
            wsWireConfReport.DisplayOptions.FrozenPaneSettings.FrozenRows = 2;

            ultraGridExcelWireConfReport.Export(wcrGrid, wsWireConfReport);
            wb1.Save(FullPathandFileName);
        }
        /// <summary>
        /// Create individual excel reports for each client branch and their subordinates.
        /// </summary>
        /// <param name="ultragridExcel"></param>
        /// <param name="mainList"></param>
        /// <param name="reportDate"></param>
        /// <param name="isMap">flag to determine if report is for Map(true) or Beken(false)</param>
        public void GenerateClientExcelReport(UltraGrid ultragridExcel, string reportDate, bool isMap)
        {
            ExcelClientReportObject excelObject = new ExcelClientReportObject(ultragridExcel, reportDate, isMap);
            progressBarPanel = new FloatingProgressBar("Client Excel Report Generation", ToExcelList.Count);
            bgwExportPDF.RunWorkerAsync(excelObject);
            progressBarPanel.ShowDialog();
        }

        private bool ExportToPDF(string directory, string filename)
        {            
            string ExcelFile = directory + "\\" + filename + ".xls";
            string PDFfile = directory + "\\" + filename + ".pdf";
            bool rvalue = false;
            //Create COM Objects
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            object unknownType = Type.Missing;

            try
            {
                //Create new instance of Excell
                //Open excel application in hidden mode
                excelApplication = new Microsoft.Office.Interop.Excel.Application
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false
                };

                //Open Excel file that needs to be converted
                if (excelApplication != null)
                {
                    excelWorkbook = excelApplication.Workbooks.Open(ExcelFile, unknownType, unknownType,
                        unknownType, unknownType, unknownType,
                        unknownType, unknownType, unknownType,
                        unknownType, unknownType, unknownType,
                        unknownType, unknownType, unknownType
                    );
                }

                //Export Excel as PDF
                //Call Excel's native export funtion (valid in Office 2007 and Office 2010)
                excelWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                    PDFfile, unknownType, unknownType, unknownType, unknownType, unknownType,
                    unknownType, unknownType);

                //if we reach here, we were able to export the excel file to pdf.
                rvalue = true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Export to PDF Error", System.Windows.Forms.MessageBoxButtons.OK);
                rvalue = false;
            }
            //Quite Ecel and release the ApplicationClass Object
            finally
            {
                // Close the workbook, quit the Excel, and clean up regardless of the results...
                if (excelWorkbook != null)
                    excelWorkbook.Close(unknownType, unknownType, unknownType);
                if (excelApplication != null)
                    excelApplication.Quit();

                releaseObject(excelWorkbook);
                releaseObject(excelApplication);
            }
            return rvalue;
        }

        public void releaseObject(object obj)
        {
            //Console.WriteLine("MethodName: releaseObject of Class: Util started");
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }

            catch (Exception exReleaseObject)
            {
                obj = null;
                //   Console.WriteLine(CMSResourceFile.REALESE_FAILED+ exReleaseObject);

            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            //Console.WriteLine("MethodName: releaseObject of Class: Util ended");
        }

        /// <summary>
        /// Save the Emails into the user's Outlook Draft folder.
        /// </summary>
        public async void SaveEmailDraft()
        {
            //perform Graph Authentication before luanching the background thread.
            if(graphHelper == null)
                graphHelper = new GraphHelper();
            graphHelper.InitializeGraph();
            await graphHelper.GreetUserAsync();

            if (graphHelper.getEmail() == null || graphHelper.getEmail().Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("Unable to login to your Microsoft Cloud account!", "Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            else
                mailUtil.SaveDraftEmail(graphHelper);

        }

        public void ResetEmailDraftList()
        {
            mailUtil.ResetList();
        }

        public bool isEmailDraftListEmpty()
        {
            return mailUtil.isListEmpty();
        }

        #region Background Worker Export PDF

        void bgwExportPDF_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarPanel.UpdateProgressBar(e.ProgressPercentage);
        }


        void bgwExportPDF_DoWork(object sender, DoWorkEventArgs e)
        {
            ExcelClientReportObject myReportObject = (ExcelClientReportObject)e.Argument;
            //        MailMessage emailDraft;
            ClientBranch lastBranch = new ClientBranch();
            System.Windows.Forms.BindingSource bindingSourceBranch = new System.Windows.Forms.BindingSource(); ;
            //VS3611 Change logic around so that we cloan the List so that the changes made in here does not affect 
            //the list for other tables...

            List<ClientBranch> masterList = new List<ClientBranch>();
            foreach (ClientBranch sourceBranch in ToExcelList)
            {
                masterList.Add(sourceBranch.Clone());
            }

            UltraGrid clientReportUltraGrid = myReportObject.ExcelGrid;
            //VS4142 determine if this is a Gift or Prepaid..
            string filename = "";
            string prefix = "";
            if (myReportObject.ReportDate.Contains("Gift"))
            {
                prefix = "Gift";
            }
            else
                prefix = "Prepaid";

            //string filename = UserSettings.Current.ExcelFileNamePrefix;
            int excelSummaryType = 0;
            string path = "";
            wb = null;

            int counter = 0; //used to count status updates.

            //Check to make sure we have something in the grid...
            if (masterList.Count > 0)
            {
                //VS4591 we have something in the grid so lets create a progrss dialog box.
                //loop through the client branches in the grid and get the names of the client branch and determine
                foreach (ClientBranch mBranch in masterList)
                {
                    counter++;
                    //get the branch name from the Prepaid Total Summary client branch...
                    if (mBranch.Category == ClientBranch.ClientBranch_Category.PREPAID_TOTAL_SUMMARY)
                    {

                        //need to save off the workbook when we encounter the next client branch Prepaid total summary category...
                        //VS3605 -- only save if there are transaction infor in the branch.
                        if (wb != null && _singleBranchReportList.Count > 0)
                        {
                            bindingSourceBranch.DataSource = _singleBranchReportList;
                            bindingSourceBranch.ResetBindings(false);
                            clientReportUltraGrid.DataSource = bindingSourceBranch;
                            ultraGridExcelClientReportExporter.Export(clientReportUltraGrid, ws, 4, 0);
                            ws.DisplayOptions.ShowGridlines = false;
                            //VI3598N -- set the scaling factory so width of report fit on a single page...
                            ws.PrintOptions.ScalingType = ScalingType.UseScalingFactor;
                            ws.PrintOptions.ScalingFactor = 70;
                            //end VI3598N
                            path = this.ExcelDirectory + "\\" + filename + ".xls";
                            wb.Save(path);
                            //Convert to PDF
                            if (ExportToPDF(this.ExcelDirectory, filename))
                            {
                                //for testing purposes only
                                //mBranch.EmailAddress = "mikec@crystalpoint.com";
                                
                                mailUtil.CreateDraftEmail(/*mBranch.EmailAddress*/lastBranch.EmailAddress, "Client Report: " + filename, "", this.ExcelDirectory + "\\" + filename + ".pdf");
                            }
                            bgwExportPDF.ReportProgress(counter);
                            _singleBranchReportList.Clear(); //clear out the contents so that we can reuse for next branch.
                        }

                        //start a new book...
                        wb = new Workbook();
                        ws = wb.Worksheets.Add("Report");

                        //manipulate the name so that it does not contains space, the word "Global" and the chars (* and _)
                        //VS4142
                        filename = prefix;
                        filename += ParseBranchName(mBranch.Name);
                        string date = ParseDate(myReportObject.ReportDate);
                        //Check to see if the mBranch.Name ends with a digit or not. If so then we need to append
                        //an underscore before the date for readability.
                        if (isDigit(filename[filename.Length - 1]))
                            filename += "_" + date;
                        else
                            filename += date;

                        CreateExcelTemplate(ref wb, ref ws, myReportObject.ReportDate, myReportObject.IsMap);

                        if (mBranch.ExcelSummary == ClientBranch.ExcelSummaryType.ALL)
                        {
                            excelSummaryType = 0;
                        }
                        else if (mBranch.ExcelSummary == ClientBranch.ExcelSummaryType.GPR_Gift_Summary_Prepaid_Total)
                        {
                            excelSummaryType = 1;
                        }
                        else //Excel summary is GPR_GIFT_Details...
                        {
                            excelSummaryType = 2;
                        }
                    }

                    //Need to remove unwanted Transaction type from the clientbranch...
                    ClientBranch adjustedBranch = ParseOutTransactionTypes(mBranch);


                    //add the branches to the singleBranchReportList so that we can construct custom report for individual clients that consist of
                    //their specific report type that they want regarding all their branches.
                    switch (excelSummaryType)
                    {
                        case 0: //all types are added (GPR Details, Gift Details, GPR Summary, Gift Summary, and Prepaid Total)  
                            if (adjustedBranch.TransactionList.Count > 0)
                                _singleBranchReportList.Add(adjustedBranch);
                            break;
                        case 1://GPR Summary, Gift Summary, and Prepaid Total
                            if (mBranch.Category == ClientBranch.ClientBranch_Category.GPR_SUMMARY || mBranch.Category == ClientBranch.ClientBranch_Category.GIFT_SUMMARY || mBranch.Category == ClientBranch.ClientBranch_Category.PREPAID_TOTAL_SUMMARY)
                            {
                                if (adjustedBranch.TransactionList.Count > 0)
                                    _singleBranchReportList.Add(adjustedBranch);
                            }
                            break;
                        case 2://GPR Detials and Gift Details
                            if (mBranch.Category == ClientBranch.ClientBranch_Category.GPR_DETAIL || mBranch.Category == ClientBranch.ClientBranch_Category.GIFT_DETAIL)
                            {
                                if (adjustedBranch.TransactionList.Count > 0)
                                    _singleBranchReportList.Add(adjustedBranch);

                            }
                            break;
                        default:
                            break;

                    }
                    lastBranch = mBranch;
                }
            }

            //need to save the last client branch workbook...
            //VS3605 if there are transaction info in there then save the workbook..
            if (wb != null && _singleBranchReportList.Count > 0)
            {
                bindingSourceBranch.DataSource = _singleBranchReportList;
                bindingSourceBranch.ResetBindings(false);
                clientReportUltraGrid.DataSource = bindingSourceBranch;
                ultraGridExcelClientReportExporter.Export(clientReportUltraGrid, ws, 4, 0);
                ws.DisplayOptions.ShowGridlines = false;
                //VI3598N -- set the scaling factory so width of report fit on a single page...
                ws.PrintOptions.ScalingType = ScalingType.UseScalingFactor;
                ws.PrintOptions.ScalingFactor = 70;
                //end VI3598N
                path = this.ExcelDirectory + "\\" + filename + ".xls";
                wb.Save(path);
                //Convert to PDF
                if (ExportToPDF(this.ExcelDirectory, filename))
                {
                    //for testing purposes only
                    //lastBranch.EmailAddress = "mikec@crystalpoint.com";
                    
                    mailUtil.CreateDraftEmail(lastBranch.EmailAddress, "Client Report: " + filename, "", this.ExcelDirectory + "\\" + filename + ".pdf");
                }
                bgwExportPDF.ReportProgress(counter);
                _singleBranchReportList.Clear(); //clear out the contents so that we can reuse for next branch.
            }
            else
                wb = null;
        }

        void bgwExportPDF_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (progressBarPanel != null)
                progressBarPanel.BeginInvoke(new Action(() => progressBarPanel.Close()));
        }

        #endregion

        /// <summary>
        /// Simple method to determine if a char is a digit or not.
        /// </summary>
        /// <param name="chr"></param>
        /// <returns></returns>
        /// 
        private bool isDigit(char chr)
        {
            if (chr >= '0' && chr <= '9')
                return true;
            return false;
        }

        /// <summary>
        /// The individual client reports have a specific formatting. This formatting results in the REMOVAL of the following transaction types:
        ///     Unloads/FI Funds Transfer
        ///     Reloads/By-Pass
        ///	    Reloads/Merchant ACQ
        ///     Unloads/Merchant ACQ
        ///     Unloads/By-Pass
        ///1/27/2020 Add the following to the list to remove from client reports:
        ///     RELOADS/MONEY TSFR RCVD
        ///     ACH DIRECT DEPOSIT 
        /// In addition to the removal of the above transaction types, we also need to swap the order of listing around for:
        /// MANUAL ADJUSTMENTS and
        /// TOTAL LOAD/UNLOAD ACTIVITY
        /// and to sum up the TransactionAmount into TOTAL LOAD/UNLOAD ACTIVITY to include MANUAL ADJUSTMENT.
        /// </summary>
        /// <param name="mBranch"></param>
        private ClientBranch ParseOutTransactionTypes(ClientBranch mBranch)
        {
            ClientBranch adjustedBranch = mBranch;
            //remove unwanted transactiontypes from the TransactionList within the ClientBranch object...
            //calculate the running total amount that we removed so that we can adjust the Total Load/Unload value...
            decimal adjustedTransct = 0;
            for(int i = 0; i < mBranch.TransactionList.Count; i++)
            {
                TransactionType mType = mBranch.TransactionList[i];

                if (mType.Transaction == TransactionType.TransactionOption.UNLOADS_FI_FUNDS_TRANSFER ||
                    mType.Transaction == TransactionType.TransactionOption.RELOADS_BYPASS ||
                    mType.Transaction == TransactionType.TransactionOption.RELOADS_MERCHANT_ACQ ||
                    mType.Transaction == TransactionType.TransactionOption.UNLOADS_MERCHANT_ACQ ||
                    mType.Transaction == TransactionType.TransactionOption.UNLOADS_BYPASS ||
                    mType.Transaction == TransactionType.TransactionOption.LOADS_MERCH_POS_FUNDING ||   //VS4070 Fix
                    mType.Transaction == TransactionType.TransactionOption.LOADS_BYPASS ||              //VS4637
                    mType.Transaction == TransactionType.TransactionOption.LOADS_MERCHANT_ACQ ||        //VS4637
                    mType.Transaction == TransactionType.TransactionOption.ACH_DIRECT_DEPOSIT ||        //VS6323
                    mType.Transaction == TransactionType.TransactionOption.RELOADS_MONEY_TRSFR_RCVD)  //VS6323
                {
                    adjustedBranch.TransactionList.Remove(mType);
                    i--; //if we remove an item from the list we need to adjust the index to compensate for that.
                    adjustedTransct+= mType.TransactionCount;
                }
            }

            //After removing the unwanted transactiontypes need to check to see if all we have left is the  Total Load/Unload transaction.
            //If so then we need to remove it because it doesn't make sense to have that entry in there when there are no transactions.
            if (adjustedBranch.TransactionList.Count == 1)
            {
                if (adjustedBranch.TransactionList[0].Transaction == TransactionType.TransactionOption.TOTAL_LOAD_UNLOAD_ACTIVITY)
                    adjustedBranch.TransactionList.Clear();
            }

            //need to swap two transaction types around. In regular reports the transactions are always listed in the following format:
            //
            //TOTAL LOAD/UNLOAD ACTIVITY
            //MANUAL ADJUSTMENTS
            //
            //Now MAP would like to swap those around so that it appears as:
            //
            //MANUAL ADJUSTMENTS
            //TOTAL LOAD/UNLOAD ACTIVITY
            //Then sum up the value for the TransactionAmount the TOTAL LOAD/UNLOAD ACTIVITY row to inlcude the MANUAL ADJUSTMENTS and the remaining transaction types
            TransactionType tmpManAdj = null;
            TransactionType tmpTotalLoadUnloadAct = null;
            decimal totalLoadUnloadSum = 0;
            for(int i = 0; i < adjustedBranch.TransactionList.Count; i++)
            {
                TransactionType mType = adjustedBranch.TransactionList[i];
                //VS4637 recalculate  Total Load Unload Activity by adding the values for Loads FI Funds Transer, Reloads FI Funds Transfer, and Manual Adjustments
                if (mType.Transaction == TransactionType.TransactionOption.LOADS_FI_FUNDS_TRANSFER || mType.Transaction == TransactionType.TransactionOption.RELOADS_FI_FUNDS_TRANSFER)
                {
                    totalLoadUnloadSum += mType.TransactionAmount;
                }
                
                if (mType.Transaction == TransactionType.TransactionOption.TOTAL_LOAD_UNLOAD_ACTIVITY)
                {

                    //adjust the transactioncount...
                    mType.TransactionCount -= adjustedTransct;

                    tmpTotalLoadUnloadAct = mType;
                }

                if (mType.Transaction == TransactionType.TransactionOption.MANUAL_ADJUSTMENT)
                {
                    tmpManAdj = mType;

                    //add / subtract manual adjustment to the adjustedAmount...
                    if (mType.DBCR1.Equals("CR"))
                        totalLoadUnloadSum += mType.TransactionAmount;
                    else
                        totalLoadUnloadSum -= mType.TransactionAmount;
                }
            }

            //VS4637 Store the recalculated values for Total Loads unloads into the transaction.
            if (tmpTotalLoadUnloadAct != null)
            {
                tmpTotalLoadUnloadAct.TransactionAmount = totalLoadUnloadSum;
                //VS6323 - If a manual adjustment causes the *TOTAL LOAD/UNLOAD ACTIVITY value to be negative we need
                //to remove the negative sign and then change the CR to DR value.
                if(tmpTotalLoadUnloadAct.TransactionAmount < 0)
                {
                    tmpTotalLoadUnloadAct.TransactionAmount = tmpTotalLoadUnloadAct.TransactionAmount * -1;
                    tmpTotalLoadUnloadAct.DBCR1 = "DR";
                }
            }

            if (tmpManAdj != null && tmpTotalLoadUnloadAct != null)
            {
                //remove this transaction if there is a swap going on...
                adjustedBranch.TransactionList.Remove(tmpManAdj);
                adjustedBranch.TransactionList.Remove(tmpTotalLoadUnloadAct);

                //add the transaction types back into the client branch object
                adjustedBranch.TransactionList.Add(tmpManAdj);
                adjustedBranch.TransactionList.Add(tmpTotalLoadUnloadAct);
            }

            //adjustment made to tally up the TransactionCount for ClientBranch category Prepaid Total, GPR Summary and Gift Summary
            if (adjustedBranch.Category == ClientBranch.ClientBranch_Category.PREPAID_TOTAL_SUMMARY ||
               adjustedBranch.Category == ClientBranch.ClientBranch_Category.GPR_SUMMARY ||
               adjustedBranch.Category == ClientBranch.ClientBranch_Category.GIFT_SUMMARY)
            {
                //recalculate the transaction counts to make sure they add up correctly...
                adjustedTransct = 0; //reset the value.
                foreach (TransactionType mType in adjustedBranch.TransactionList)
                {
                    if (mType.Transaction == TransactionType.TransactionOption.LOADS_FI_FUNDS_TRANSFER ||
                       mType.Transaction == TransactionType.TransactionOption.RELOADS_FI_FUNDS_TRANSFER)
                    {
                        adjustedTransct += mType.TransactionCount;
                    }

                    //update the value into the client branch...
                    if (mType.Transaction == TransactionType.TransactionOption.TOTAL_LOAD_UNLOAD_ACTIVITY)
                    {
                        mType.TransactionCount = adjustedTransct;
                    }
                }
            }

            return adjustedBranch;
        }

        /// <summary>
        /// Parse the BranchName so that the new name does not contain any spaces, asterisk, underscore, and the word Global.
        /// This new name will be used as part of a report filename.
        /// </summary>
        /// <param name="inName"></param>
        /// <returns></returns>
        private String ParseBranchName(String inName)
        {
            string modifiedName = inName;

            //parse out the unwanted chars...
            //VS4583 Beken reports. Test report uses CU name for private bank: Beken Bus Dev 1. Need to include the digits in the
            //reg pattern.
            Regex pattern = new Regex("([A-Za-z1-9]+)");
            MatchCollection mCollection = pattern.Matches(inName);
            if (mCollection.Count > 0) //reset the modifiedName so that we can add the parsed out content to it...
            {
                modifiedName = "";
                for (int i = 0; i < mCollection.Count; i++)
                {
                    modifiedName += mCollection[i].Value;
                }
            }
            if(modifiedName.EndsWith("Global", StringComparison.CurrentCultureIgnoreCase))
            {
                //parse out the word global which is lenght of 6 chars...
                modifiedName = modifiedName.Substring(0, modifiedName.Length - 6);
            }
            return modifiedName;
        }

        /// <summary>
        /// Parse out the backslashes from the date string so that we can use the value as part of the filename.
        /// </summary>
        /// <param name="inDate"></param>
        /// <returns></returns>
        private String ParseDate(string inDate)
        {
            string modifiedDate = inDate;

            //parse out the backslashes...
            Regex pattern = new Regex("([0-9]+/[0-9]+/[0-9]+)");
            MatchCollection mCollection = pattern.Matches(inDate);
            if (mCollection.Count > 0)
            {
                modifiedDate = "";
                int dateLength = mCollection[0].Value.Length;
                //we have a match for the date pattern
                //now replace the / values with _ (10/13/2010 --> 10_13_2010
                pattern = new Regex("([0-9]+)");
                mCollection = pattern.Matches(mCollection[0].Value);
                for (int i = 0; i < mCollection.Count; i++)
                {
                    modifiedDate += mCollection[i].Value;
                }
            }
            return modifiedDate;
        }

        /// <summary>
        /// Each individual client report needs to have display the MAP corporate logo and address. This method builds this
        /// template into the worksheet.
        /// </summary>
        /// <param name="wb">Excel workbook</param>
        /// <param name="ws">Excel worksheet</param>
        /// <param name="reportDate">Date that report was generated</param>
        /// <param name="isMap">bool value to determine report type</param>
        /// <returns></returns>
        private bool CreateExcelTemplate(ref Workbook wb, ref Worksheet ws, String reportDate, bool isMap)
        {
            bool templateCreated = true;
            string previousDate = "";

            try
            {
                //VS4170 parse the date so that we only have the date and nothing else...
                Regex pattern = new Regex("([0-9]+/[0-9]+/[0-9]+)");
                MatchCollection mCollection = pattern.Matches(reportDate);
                reportDate = mCollection[0].Value;
                
                DateTime dt = Convert.ToDateTime(reportDate);

                //calculate previous date by subtracting 1 day from the report date.
                dt = dt.AddDays(-1);
                previousDate = dt.ToString("MM/dd/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo);
            }
            catch (Exception e1) 
            {
                previousDate = reportDate;
            };

            try
            {
                //Create the cell formatting...
                IWorksheetCellFormat cellFormat = wb.CreateNewWorksheetCellFormat();
                cellFormat.Font.Name = "Arial";
                cellFormat.Font.Bold = ExcelDefaultableBoolean.True;
                cellFormat.Font.Color = Color.Black;
                //change font point from 24 to 23 (twip it is 480 to 460) so that the pagebreak view in excel will fit on one page...
                //VS3606
                cellFormat.Font.Height = 460;
                ws.Rows[0].Cells[0].CellFormat.SetFormatting(cellFormat);
                //VS4583 added Beken reports
                if(isMap)
                    ws.Rows[0].Cells[0].Value = "Member Access Processing GPR and Gift Card Settlements";
                else
                    ws.Rows[0].Cells[0].Value = "BEKEN GPR and Gift Card Settlements";
                ws.Rows[0].Height = 570;

                IWorksheetCellFormat cellFormat2 = wb.CreateNewWorksheetCellFormat();
                cellFormat2.Font.Name = "Arial";
                cellFormat2.Font.Color = Color.Black;
                cellFormat2.Font.Bold = ExcelDefaultableBoolean.False;
                cellFormat2.Font.Height = 240;
                ws.Rows[1].Cells[0].CellFormat.SetFormatting(cellFormat2);

                //ws.Rows[1].Cells[0].Value = "16000 Christensen Rd. Suite 200, Tukwila WA 98188";
                //VS6323 Change address to new value.
                ws.Rows[1].Cells[0].Value = "Member Access Processing, LLC - 20829 72nd Ave S, Ste 600, Kent WA 98032  ";
                ws.Rows[1].Height = 345;

                ws.Rows[2].Cells[0].CellFormat.SetFormatting(cellFormat2);
                ws.Rows[2].Cells[0].Value = "Tel: 866-598-0698   Fax: 206-439-0045";
                ws.Rows[2].Height = 435;

                Image image;
                if (isMap)
                {
                    //image = MAPacificReportUtility.Properties.Resources.logo2;
                    //VS6323 change the logo to be the newer version sent over from MAP
                    image = MAPacificReportUtility.Properties.Resources.map1sm;
                }
                else
                    image = MAPacificReportUtility.Properties.Resources.Beken;
                WorksheetImage imageShape = new WorksheetImage(image);
                
                imageShape.TopLeftCornerCell = ws.Rows[2].Cells[3];  
                //change the cell column for the new updated MAP logo
                //ws.Rows[1].Cells[3];
                //if(isMap)
                //    imageShape.BottomRightCornerCell = ws.Rows[3].Cells[4];
                //else //logo for Beken.bmp looks squished so we need to make the right corner one cell over.
                //4,5
                imageShape.BottomRightCornerCell = ws.Rows[5].Cells[6];

                imageShape.BottomRightCornerPosition = new PointF(100, 100);

                ws.Shapes.Add(imageShape);

                IWorksheetCellFormat cellFormat3 = wb.CreateNewWorksheetCellFormat();
                cellFormat3.Font.Name = "Arial";
                cellFormat3.Font.Bold = ExcelDefaultableBoolean.False;
                cellFormat3.Font.Height = 240;
                cellFormat3.Font.Color = Color.Red;
                ws.Rows[3].Cells[0].CellFormat.Alignment = HorizontalCellAlignment.Right; 
                ws.Rows[3].Cells[0].Value = "Settlement Date: ";
                ws.Rows[3].Cells[1].CellFormat.SetFormatting(cellFormat3);
                ws.Rows[3].Cells[1].Value = previousDate;
                //VS3607 -- increase the row height for the date.
                ws.Rows[3].Height = 550;

                //ws.Rows[3].Cells[1].CellFormat.SetFormatting(cellFormat2);
                //ws.Rows[3].Cells[1].Value = " <-Previous Business Day";
                //ws.Rows[3].Height = 550;

                //set the width for the first 5 columns
                ws.Columns[0].Width = 3321; //90 pixels
                ws.Columns[1].Width = 10110;//276 pixels
                ws.Columns[2].Width = 5941; //161 pixels
                ws.Columns[3].Width = 6236; //170 pixels


            }catch(Exception exc)
            {
                templateCreated = false;
            }
            return templateCreated;
        }

        private void ExtractFileName(string inName)
        {
            //create the filename for the excel...
            //value in textBoxReportDate is default to current date: PrepaidImport(12/8/2010)
            //However users can type whatever value they want and also if current date does not match the runtime date then a warning
            //message is appended into that textbox.

            string fileName = inName;
            Regex pattern = new Regex("([0-9]+/[0-9]+/[0-9]+)");
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
                    ExcelFileName = fileName;
                }
                else if (index > 0 && fileName.StartsWith("GiftImport"))
                {
                    fileName = fileName.Substring(0, index);
                    fileName += date + ")";
                    ExcelFileName = fileName;
                }
                else
                {
                    //date happens to sit in the middle of the file name...
                    //"FileName 12/25/2010 Quarter4"
                    if (index > 0)
                    {
                        fileName = inName.Substring(0, index);
                        fileName += "Import(";
                        fileName += date;
                        fileName += inName.Substring(index + dateLength);
                        fileName += ")";
                        ExcelFileName = fileName;
                    }
                    else //date happens to be at the start of the filename...
                    {
                        fileName = date;
                        ExcelFileName = fileName;
                    }
                }
            }
            else
            {
                ExcelFileName = inName;
            }
        }

        #region inner class ExcelClientReportObject
        /// <summary>
        /// Inner class to help facilitate the passing of information between seperate objects to determine the
        /// report date and is report is a MAP report or Beken report. Also the string in the report date
        /// helps to determine if report is Gift or Prepaid.
        /// </summary>
        class ExcelClientReportObject
        {
            //properties
            private UltraGrid excelGrid;
            public UltraGrid ExcelGrid
            {
                get { return excelGrid; }
            }

            private string reportDate;
            public string ReportDate
            {
                get { return reportDate; }
            }

            private bool isMap;
            public bool IsMap
            {
                get { return isMap; }
            }

            public ExcelClientReportObject(UltraGrid ultragridExcel, string inDate, bool Map)
            {
                excelGrid = ultragridExcel;
                reportDate = inDate;
                isMap = Map;
            }
        }
        #endregion
    }
}

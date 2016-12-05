using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using Rock;
using Rock.Extension;
using Rock.Model;
using Rock.Web.UI;
using Rock.Data;
using Rock.Web;

namespace RockWeb.Plugins.com_shepherdchurch.v5_financials
{
    [DisplayName( "Export to GL" )]
    [Category( "com_shepherdchurch > V5 Financials" )]
    [Description( "Export the current batch to the GL system." )]

    public partial class ExportToGL : RockBlock, ISecondaryBlock
    {
        FinancialBatch _batch = null;
        protected int IsExported = 0;

        #region Base Method Overrides

        /// <summary>
        /// Raises the <see cref="E:System.Web.UI.Control.Init" /> event.
        /// </summary>
        /// <param name="e">An <see cref="T:System.EventArgs" /> object that contains the event data.</param>
        protected override void OnInit( EventArgs e )
        {
            base.OnInit( e );

            // this event gets fired after block settings are updated. it's nice to repaint the screen if these settings would alter it
            this.BlockUpdated += Block_BlockUpdated;
            this.AddConfigurationUpdateTrigger( upnlContent );
        }

        /// <summary>
        /// Initialize basic information about the page structure and setup the default content.
        /// </summary>
        /// <param name="sender">Object that is generating this event.</param>
        /// <param name="e">Arguments that describe this event.</param>
        protected void Page_Load( object sender, EventArgs e )
        {
            ScriptManager.GetCurrent( this.Page ).RegisterPostBackControl( lbDownload );

            _batch = new FinancialBatchService( new RockContext() ).Get( 26867 );

            if ( !string.IsNullOrWhiteSpace( PageParameter( "batchId" ) ) )
            {
                _batch = new FinancialBatchService( new RockContext() ).Get( PageParameter( "batchId" ).AsInteger() );
            }

            if ( _batch != null )
            {
                _batch.LoadAttributes();
                IsExported = (_batch.GetAttributeValue( "GLExported" ).AsBoolean() == true ? 1 : 0);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Get the GLRecords for the given batch, with the appropriate information about
        /// the export data.
        /// </summary>
        /// <param name="batch">Batch to be exported.</param>
        /// <param name="date">The date of this deposit.</param>
        /// <param name="accountingPeriod">Accounting period as defined in the GL system.</param>
        /// <param name="journalType">The type of journal entry to create as defined in the GL system.</param>
        /// <returns>A collection of GLRecord objects to be imported into the GL system.</returns>
        List<GLRecord> GLRecordsForBatch( FinancialBatch batch, DateTime date, string accountingPeriod, string journalType )
        {
            List<GLRecord> records = new List<GLRecord>();

            //
            // Load all the transaction details, load their attributes and then group
            // by the account attributes, GLBankAccount+GLCompany+GLFund.
            //
            var transactions = batch.Transactions
                .SelectMany( t => t.TransactionDetails )
                .ToList();
            foreach ( var d in transactions )
            {
                d.LoadAttributes();
                d.Account.LoadAttributes();
            }
            var accounts = transactions.GroupBy( d => new { GLBankAccount = d.Account.GetAttributeValue( "GLBankAccount" ), GLCompany = d.Account.GetAttributeValue( "GLCompany" ), GLFund = d.Account.GetAttributeValue( "GLFund" ) }, d => d ).OrderBy( g => g.Key.GLBankAccount );

            //
            // Go through each group and build the line items.
            //
            foreach ( var grp in accounts )
            {
                GLRecord record = new GLRecord();

                //
                // Build the bank account deposit line item.
                //
                record.AccountingPeriod = accountingPeriod;
                record.AccountNumber = grp.Key.GLBankAccount;
                record.Amount = grp.Sum( d => d.Amount );
                record.Company = grp.Key.GLCompany;
                record.Date = date;
                record.Department = string.Empty;
                record.Description1 = batch.Name + " (" + batch.Id.ToString() + ")";
                record.Description2 = string.Empty;
                record.Fund = grp.Key.GLFund;
                record.Journal = "0";
                record.JournalType = journalType;
                record.Project = string.Empty;

                records.Add( record );

                //
                // Build each of the revenue fund withdrawls.
                //
                foreach ( var grpTransactions in grp.GroupBy(t => t.AccountId, t => t ) )
                {
                    record = new GLRecord();

                    record.AccountingPeriod = accountingPeriod;
                    record.AccountNumber = grpTransactions.First().Account.GetAttributeValue( "GLRevenueAccount" );
                    record.Amount = -(grpTransactions.Sum( t => t.Amount ));
                    record.Company = grp.Key.GLCompany;
                    record.Date = date;
                    record.Department = grpTransactions.First().Account.GetAttributeValue( "GLRevenueDepartment" );
                    record.Description1 = grpTransactions.First().Account.Name;
                    record.Description2 = string.Empty;
                    record.Fund = grp.Key.GLFund;
                    record.Journal = "0";
                    record.JournalType = journalType;
                    record.Project = string.Empty;

                    records.Add( record );
                }
            }

            return records;
        }

        /// <summary>
        /// Hook so that other blocks can set the visibility of all ISecondaryBlocks on its page
        /// </summary>
        /// <param name="visible">if set to <c>true</c> [visible].</param>
        public void SetVisible( bool visible )
        {
            pnlMain.Visible = visible;
        }

        #endregion

        #region Event Handlers

        /// <summary>
        /// Handles the BlockUpdated event of the control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected void Block_BlockUpdated( object sender, EventArgs e )
        {
        }

        /// <summary>
        /// Handles the Click event of the lbShowExport control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected void lbShowExport_Click( object sender, EventArgs e )
        {
            dpDate.SelectedDate = RockDateTime.Now;
            tbAccountingPeriod.Text = GetUserPreference( "com.shepherdchurch.exporttogl.accountingperiod" );
            tbJournalType.Text = GetUserPreference( "com.shepherdchurch.exporttogl.journaltype" );
            nbAlreadyExported.Visible = _batch.GetAttributeValue( "GLExported" ).AsBoolean();

            pnlExportModal.Visible = true;
            mdExport.Show();
        }

        /// <summary>
        /// Handles the Click event of the lbSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected void lbExportSave_Click( object sender, EventArgs e )
        {
            SetUserPreference( "com.shepherdchurch.exporttogl.accountingperiod", tbAccountingPeriod.Text );
            SetUserPreference( "com.shepherdchurch.exporttogl.journaltype", tbJournalType.Text );

            mdExport.Hide();
            pnlExportModal.Visible = false;

            //
            // After the page updates, simulate a click on the download link, wait 1
            // second and then reload the page (non-postback) so that the UI will update
            // to reflect changes about the batch.
            //
            string script = string.Format( "document.getElementById('{0}').click(); setTimeout(function() {{ location.reload(false); }}, 1000);", lbDownload.ClientID );
            ScriptManager.RegisterStartupScript( Page, Page.GetType(), "PerformExport", script, true );
        }

        /// <summary>
        /// Handles the Click event of the lbDownload control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected void lbDownload_Click( object sender, EventArgs e )
        {
            var parameters = RockPage.PageParameters();

            var records = GLRecordsForBatch( _batch, dpDate.SelectedDate.Value, tbAccountingPeriod.Text.Trim(), tbJournalType.Text.Trim() );

            //
            // Update the batch to reflect that it has been exported.
            //
            using ( var rockContext = new RockContext() )
            {
                FinancialBatch batch = new FinancialBatchService( rockContext ).Get( _batch.Id );

                batch.LoadAttributes();
                batch.SetAttributeValue( "GLExported", "true" );
                batch.SaveAttributeValues( rockContext );
                IsExported = 1;

                rockContext.SaveChanges();
            }

            //
            // Send the results as a CSV file for download.
            //
            Page.EnableViewState = false;
            Page.Response.Clear();
            Page.Response.ContentType = "text/csv";
            Page.Response.AppendHeader( "Content-Disposition", "attachment; filename=GLTRN2000.txt" );
            Page.Response.Write( string.Join( "\r\n", records.Select( r => r.ToString() ).ToArray() ) );
            Page.Response.Flush();
            Page.Response.End();
        }

        #endregion
    }

    class GLRecord
    {
        public string Company { get; set; }
        public string Fund { get; set; }
        public string AccountingPeriod { get; set; }
        public string JournalType { get; set; }
        public string Journal { get; set; }

        public DateTime Date { get; set; }

        public string Description1 { get; set; }

        public string Description2 { get; set; }

        public string Department { get; set; }
        public string AccountNumber { get; set; }

        public decimal Amount { get; set; }

        public string Project { get; set; }

        public override string ToString()
        {
            return string.Format( "\"00000\",\"0{0}00{1}{2}{3}{4}\",\"000\",\"{5}\",\"{6}\",\"{7}\",\"{8}{9}\",\"{10}\",\"{11}\"",
                (Company ?? string.Empty).PadLeft(3, '0').TrimLength( 3),
                (Fund ?? string.Empty).PadLeft(3, '0').TrimLength( 3),
                (AccountingPeriod ?? string.Empty).PadLeft(2, '0').TrimLength( 2),
                (JournalType ?? string.Empty).PadLeft(2, '0').TrimLength( 2),
                (Journal ?? string.Empty).PadLeft(5, '0').TrimLength( 5),
                Date.ToString("MMddyy"),
                (Description1 ?? string.Empty).TrimLength( 30),
                (Description2 ?? string.Empty).TrimLength( 30),
                (Department ?? string.Empty).PadLeft(3, '0').TrimLength( 3),
                (AccountNumber ?? string.Empty).PadLeft(9, '0').TrimLength( 9),
                Math.Round(Amount * 100).ToString("0"),
                (Project ?? string.Empty).TrimLength(30));
        }
    }

    static class StringExtensions
    {
        public static string TrimLength( this string value, int maxLength )
        {
            if ( string.IsNullOrEmpty( value ) )
            {
                return value;
            }

            return value.Length <= maxLength ? value : value.Substring( 0, maxLength );
        }
    }
}

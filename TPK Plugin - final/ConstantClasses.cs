namespace Audit.Kpmg.Plugins.SummaryReportCreation
{
    class JournalEntites
    {
        //
        public const string logicalname = "cr16a_fico_journalentries_headers";
        public const string name="cr16a_name";
        public const string JournalId = "cr16a_je_id";
        public const string CompanyCode = "cr719_companycode";
        public const string DocType = "cr719_documenttype";
        public const string Currency = "cr719_currencysiegwerk";
        public const string PostingPeriod = "cr16a_posting_period";
        public const string FiscalYear = "cr719_fiscalyear";
        public const string RecursionStartDate = "cr719_startdate";
        public const string RecursionEndDate = "cr719_enddate";
        public const string PostingDate = "cr16a_posting_date";
        public const string DocumnetDate = "cr16a_document_date";
        public const string TranslationDate = "cr16a_translation_date";
        public const string Reference = "cr16a_reference";
        public const string HeaderText = "cr16a_header_text";
        public const string ReversalDate = "cr16a_reversal_date";

        public const string IsExcelJournal = "cr16a_isexceljournal";

        public const string logicalnameLine = "cr16a_fico_journalentries_lineitems";
        public const string JournalLineId = "cr16a_je_id";
        public const string Headerid = "cr16a_je_header_id";
        public const string nameLine = "cr16a_name";
        public const string AccountType = "cr719_accounttype";
        public const string G_LAccount = "cr719_glaccount";
        public const string DB_CR_Account = "cr16a_debitcredit";
        public const string Amount = "cr16a_amount";
        public const string LineItem = "cr16a_line_item_text";
        public const string CostCenter = "cr719_costcenter";
        public const string COOrder = "cr719_coorder";
        public const string WBS = "cr16a_wbs_psp_element";
        public const string ValueDate = "cr16a_valuedate";
        public const string TaxCode = "cr719_taxcode";
        public const string TaxJuris_Diction = "cr719_taxjurisdiction";
        public const string Trading_partn = "cr719_tradingpartnfi";
        public const string ProfitCenter = "cr719_profitcenter";
        public const string FunctArea = "cr719_functionalarea";
        public const string Assignment = "cr16a_assignment";
        public const string TransType = "cr719_transtype";
        public const string SalesOrg = "cr719_salesorg";
        public const string Distr_chan = "cr16a_distr_chan";
        public const string Division = "cr16a_division";
        public const string Customer = "cr719_customer";
        public const string Country = "cr719_country";
        public const string Trading_Pt = "cr719_tradingptcopa";
        public const string Product = "cr16a_product";
        public const string Plant = "cr719_plant";
        public const string OriginId = "cr16a_originid";
        public const string Bus_type = "cr16a_bus_type_ictp";
        public const string Line_Comment = "cr719_linelevelcomments";
        public const string IsExcelJournalLine = "cr16a_isexcelgenerated";

    }
    public static class AnnotationEntity
    {
        public const string logicalname = "annotation";
        public const string primaryAttribute = "subject";
        public const string documentBody = "documentbody";
        public const string primaryId = "annotationid";
        public const string fileName = "filename";

    }
    public class AmountDetail
    {
        public string HeaderID { get; set; }
        public string CreditORDebit { get; set; }
        public decimal Amount { get; set; }
    }
}

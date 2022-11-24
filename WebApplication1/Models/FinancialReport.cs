using System;
using System.Collections.Generic;

namespace WebApplication1.Models
{
    public class FinancialReport
    {
        public int ReportId { get; set; }
        public Nullable<int> BatchId { get; set; }
        public Nullable<bool> Approved { get; set; }
        public Nullable<bool> PrintStatus { get; set; }
        public string FundCode { get; set; }
        public string StartPeriod { get; set; }
        public string EndPeriod { get; set; }
        public Nullable<decimal> OpeningFundValue { get; set; }
        public Nullable<decimal> FundNetContrbution { get; set; }
        public Nullable<decimal> AssessmentForAdmin { get; set; }
        public Nullable<decimal> NetInvestmentReturn { get; set; }
        public Nullable<decimal> GrantsFromFund { get; set; }
        public Nullable<decimal> TransfersToCharitableGiftFund { get; set; }
        public Nullable<decimal> ClosingValue { get; set; }
        public Nullable<decimal> OpeningBalanceGrantMoney { get; set; }
        public Nullable<decimal> OpeningUnrestrictedCapitalBalance { get; set; }
        public Nullable<decimal> ClosingBalanceGrantMoney { get; set; }
        public Nullable<decimal> ClosingUnrestrictedCapitalBalance { get; set; }
        public Nullable<decimal> TotalGlGifts { get; set; }
        public Nullable<decimal> TotalGrants { get; set; }
        public string FundName { get; set; }

        public virtual ICollection<DARContact> DARContact { get; set; }
    }
}

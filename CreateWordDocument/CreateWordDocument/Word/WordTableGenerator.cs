using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    public static class WordTableGenerator
    {
        public static WordTable GetSummaryWordTable(DocumentDataSet ws)
        {
            string date_format = @"dd-MMM-yyyy";
            string number_format = @"#,##0.00;(#,##0.00)";

            int firstColW = (int)(115.7 * 56.7), secondColW = (int)(37.6 * 56.7), thirdColW = (int)(36.5 * 56.7);

            var wordTable = new WordTable(3);
            foreach (var item in DocumentGeneratorHelper.RmaCompanyInformation)
            {
                wordTable.AddInvoiceRow(item.Item1, item.Item2);
            }
            wordTable.AddInvoiceRow(string.Format("Date:{0}", DateTime.Now.ToString(date_format)));
            wordTable.AddInvoiceRow(" ");
            wordTable.AddInvoiceRow(string.Format("To:{0}", ws.AssumingCompanyName));
            wordTable.AddInvoiceRow("Attn: Reinsurance Department");
            wordTable.AddInvoiceRow(string.Format("Re: {0} Business as of {1}", ws.AssumingCompanyName, ws.QuarterEndDate.ToString(date_format)));

            wordTable.AddInvoiceRow("Management Fees Summary", true, 1);
            wordTable.MergeCurrentRow(1, 1);
            wordTable.CurrentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
            wordTable.CurrentRow[1].FontSize = 12;
            wordTable.CurrentRow[1].Width = (secondColW + thirdColW);

            wordTable.AddInvoiceRow("Current Period", true, 1);
            wordTable.MergeCurrentRow(1, 1);
            wordTable.CurrentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
            wordTable.CurrentRow[1].Width = (secondColW + thirdColW);

            wordTable.AddInvoiceRow(" ");
            wordTable.AddInvoiceRow("ALL UW YEARS COMBINED", "Business In CAD", "Business In USD", true);
            var claimSummaries = ws.ReportDtos ?? new List<DocumentDataDto>();

            var group1 = claimSummaries.GroupBy(a => a.FeeType).OrderBy(a => a.Key);

            decimal total1 = 0.0m, total2 = 0.0m;
            foreach (var g1 in group1)
            {
                wordTable.AddInvoiceRow(" ");
                var feeTypeDesc = g1.Key == FeeTypeEnum.TotalManagementFees ? "Total Management Fees" : "Total Profit Commission & Management Fee Adjustments";

                wordTable.AddInvoiceRow(feeTypeDesc, true);

                var group2 = g1.GroupBy(a => a.UnderwritingYear).OrderBy(a => a.Key);

                decimal subTotal1 = 0.0m, subTotal2 = 0.0m;
                foreach (var g2 in group2)
                {
                    var val1 = string.Format("UW Year {0}", g2.Key.ToString());
                    var val2 = g2.Sum(a => a.BusinessInCAD).ToString(number_format);
                    var val3 = g2.Sum(a => a.BusinessInUSD).ToString(number_format);
                    wordTable.AddInvoiceRow(val1, val2, val3);
                    wordTable.CurrentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;
                    wordTable.CurrentRow[2].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;

                    subTotal1 += g2.Sum(a => a.BusinessInCAD);
                    subTotal2 += g2.Sum(a => a.BusinessInUSD);
                }
                wordTable.AddInvoiceRow(string.Empty, "-------------------------", "-------------------------");
                wordTable.AddInvoiceRow("Sub Total", subTotal1.ToString(number_format), subTotal2.ToString(number_format), true);

                total1 += subTotal1;
                total2 += subTotal2;
            }

            wordTable.AddInvoiceRow(" ");
            var totalV1 = "Total Amount Due to RMA this Quarter";
            wordTable.AddInvoiceRow(totalV1, total1.ToString(number_format), total2.ToString(number_format), true);

            //bottom
            wordTable.AddInvoiceRow(" ");
            wordTable.AddInvoiceRow("Please wire these funds to our account within 10 working days.");
            wordTable.AddInvoiceRow(" ");
            wordTable.AddInvoiceRow("Best regards,");

            wordTable.AddInvoiceRow(" ");
            wordTable.AddInvoiceRow(" ");

            wordTable.AddInvoiceRow("Anna Lee");
            wordTable.AddInvoiceRow("Associate Director,");
            wordTable.AddInvoiceRow("Life Reinsurance Administration and Financial Reporting");

            //
            wordTable.ColumnWidth = new List<int>() { firstColW, secondColW, thirdColW };
            wordTable.AdjustColumnWidth();

            return wordTable;
        }

        public static WordTable GetITDWordTable(DocumentDataSet ws)
        {
            string date_format = @"dd-MMM-yyyy";
            string number_format = @"#,##0.00;(#,##0.00)";

            int firstColW = (int)(115.7 * 56.7), secondColW = (int)(37.6 * 56.7), thirdColW = (int)(36.5 * 56.7);

            var wordTable = new WordTable(3);
            foreach (var item in DocumentGeneratorHelper.RmaCompanyInformation)
            {
                wordTable.AddInvoiceRow(item.Item1, item.Item2);
            }

            wordTable.AddInvoiceRow("Management Fees Summary", true, 1);
            wordTable.MergeCurrentRow(1, 1);
            wordTable.CurrentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
            wordTable.CurrentRow[1].FontSize = 12;
            wordTable.CurrentRow[1].Width = (secondColW + thirdColW);

            wordTable.AddInvoiceRow(string.Format("From Inception Date to {0}", ws.QuarterEndDate.ToString(date_format)), true, 1);
            wordTable.MergeCurrentRow(1, 1);
            wordTable.CurrentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
            wordTable.CurrentRow[1].Width = (secondColW + thirdColW);

            wordTable.AddInvoiceRow(" ");
            wordTable.AddInvoiceRow("  ", "Business In CAD", "Business In USD", true);
            var claimSummaries = ws.ReportITDDtos ?? new List<DocumentDataDto>();

            var group1 = claimSummaries.GroupBy(a => a.FeeType).OrderBy(a => a.Key);

            decimal total1 = 0.0m, total2 = 0.0m;
            foreach (var g1 in group1)
            {
                wordTable.AddInvoiceRow(" ");
                var feeTypeDesc = g1.Key == FeeTypeEnum.TotalManagementFees ? "Total Management Fees" : "Total Profit Commission & Management Fee Adjustments";

                wordTable.AddInvoiceRow(feeTypeDesc, true);

                var group2 = g1.GroupBy(a => a.UnderwritingYear).OrderBy(a => a.Key);

                decimal subTotal1 = 0.0m, subTotal2 = 0.0m;
                foreach (var g2 in group2)
                {
                    var val1 = string.Format("UW Year {0}", g2.Key.ToString());
                    var val2 = g2.Sum(a => a.BusinessInCAD).ToString(number_format);
                    var val3 = g2.Sum(a => a.BusinessInUSD).ToString(number_format);
                    wordTable.AddInvoiceRow(val1, val2, val3);
                    wordTable.CurrentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;
                    wordTable.CurrentRow[2].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;

                    subTotal1 += g2.Sum(a => a.BusinessInCAD);
                    subTotal2 += g2.Sum(a => a.BusinessInUSD);
                }
                wordTable.AddInvoiceRow(string.Empty, "-------------------------", "-------------------------");
                wordTable.AddInvoiceRow("Sub Total", subTotal1.ToString(number_format), subTotal2.ToString(number_format), true);

                total1 += subTotal1;
                total2 += subTotal2;
            }

            wordTable.AddInvoiceRow(" ");
            var totalV1 = "Total Amount Due to RMA from Inception to Date";
            wordTable.AddInvoiceRow(totalV1, total1.ToString(number_format), total2.ToString(number_format), true);

            //
            wordTable.ColumnWidth = new List<int>() { firstColW, secondColW, thirdColW };
            wordTable.AdjustColumnWidth();

            return wordTable;
        }

        public static WordTable GetUWYearWordTable(DocumentDataSet ws)
        {
            string date_format = @"dd-MMM-yyyy";
            string number_format = @"#,##0.00;(#,##0.00)";

            int firstColW = (int)(115.7 * 56.7), secondColW = (int)(37.6 * 56.7), thirdColW = (int)(36.5 * 56.7);

            var wordTable = new WordTable(3);
            foreach (var item in DocumentGeneratorHelper.RmaCompanyInformation)
            {
                wordTable.AddInvoiceRow(item.Item1, item.Item2);
            }

            wordTable.AddInvoiceRow("Management Fees Summary", true, 1);
            wordTable.MergeCurrentRow(1, 1);
            wordTable.CurrentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
            wordTable.CurrentRow[1].FontSize = 12;
            wordTable.CurrentRow[1].Width = (secondColW + thirdColW);

            wordTable.AddInvoiceRow("Current Period", true, 1);
            wordTable.MergeCurrentRow(1, 1);
            wordTable.CurrentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
            wordTable.CurrentRow[1].Width = (secondColW + thirdColW);

            wordTable.AddInvoiceRow(" ");
            wordTable.AddInvoiceRow(string.Format("Underwriting Year - {0}", ws.UnderwritingYear.ToString()), "Business In CAD", "Business In USD", true);
            var claimSummaries = ws.ReportDtos ?? new List<DocumentDataDto>();

            var group1 = claimSummaries.GroupBy(a => a.FeeType).OrderBy(a => a.Key);

            decimal total1 = 0.0m, total2 = 0.0m;
            foreach (var g1 in group1)
            {
                wordTable.AddInvoiceRow(" ");
                var feeTypeDesc = g1.Key == FeeTypeEnum.TotalManagementFees ? "Total Management Fees" : "Total Profit Commission & Management Fee Adjustments";

                wordTable.AddInvoiceRow(feeTypeDesc, true);

                var group2 = g1.GroupBy(a => new { a.TreatyId, a.TreatyNumber, a.TreatyName }).OrderBy(a => a.Key.TreatyName);

                decimal subTotal1 = 0.0m, subTotal2 = 0.0m;
                foreach (var g2 in group2)
                {
                    var val1 = string.Format("{0} {1}", g2.Key.TreatyNumber, g2.Key.TreatyName);
                    var val2 = g2.Sum(a => a.BusinessInCAD).ToString(number_format);
                    var val3 = g2.Sum(a => a.BusinessInUSD).ToString(number_format);
                    wordTable.AddInvoiceRow(val1, val2, val3);
                    wordTable.CurrentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;
                    wordTable.CurrentRow[2].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;

                    subTotal1 += g2.Sum(a => a.BusinessInCAD);
                    subTotal2 += g2.Sum(a => a.BusinessInUSD);
                }
                wordTable.AddInvoiceRow(string.Empty, "-------------------------", "-------------------------");
                wordTable.AddInvoiceRow("Sub Total", subTotal1.ToString(number_format), subTotal2.ToString(number_format), true);

                total1 += subTotal1;
                total2 += subTotal2;
            }

            wordTable.AddInvoiceRow(" ");
            var totalV1 = "Net Amount Due to RMA this Quarter";
            wordTable.AddInvoiceRow(totalV1, total1.ToString(number_format), total2.ToString(number_format), true);

            //
            wordTable.ColumnWidth = new List<int>() { firstColW, secondColW, thirdColW };
            wordTable.AdjustColumnWidth();

            return wordTable;
        }
    }
}

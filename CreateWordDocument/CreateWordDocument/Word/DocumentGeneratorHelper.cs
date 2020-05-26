using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    public static class DocumentGeneratorHelper
    {
        public static List<Tuple<string, bool>> RmaCompanyInformation = new List<Tuple<string, bool>>()
        {
            new Tuple<string, bool>("Reinsurance Management Associates, Inc.",true),
            new Tuple<string, bool>("170 University Ave., Suite 500",false),
            new Tuple<string, bool>( "Toronto, ON  M5H 3B3",false),
            new Tuple<string, bool>("Tel.:  (416) 408-4966",false),
            new Tuple<string, bool>("Fax:  (416) 408-2262",false)
        };

        public static void AddInvoiceRow(this WordTable table, string content, bool bold = false, int colIndex = 0)
        {
            var currentRow = table.AddRow();
            currentRow[colIndex].Content = content;
            currentRow[colIndex].FontBold = bold;

            currentRow[0].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.LEFT;
            currentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;
            currentRow[2].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;
        }

        public static void AddInvoiceRow(this WordTable table, string content1, string content2, string content3, bool bold = false)
        {
            var currentRow = table.AddRow();
            currentRow[0].Content = content1;
            currentRow[1].Content = content2;
            currentRow[2].Content = content3;

            currentRow[0].FontBold = bold;
            currentRow[1].FontBold = bold;
            currentRow[2].FontBold = bold;

            currentRow[0].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.LEFT;
            currentRow[1].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;
            currentRow[2].H_Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;
        }
    }
}

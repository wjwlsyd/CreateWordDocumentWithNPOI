using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    public class DocumentDataSet
    {
        public DocumentDataSet()
        {
         
        }
        public WorkSheetTypeEnum WorkSheetType { get; set; }
        public string WorkSheetName { get; set; }
        public string AssumingCompanyName { get; set; }
        public int? UnderwritingYear { get; set; }
        public DateTime QuarterEndDate { get; set; }

        public List<DocumentDataDto> ReportDtos { get; set; }
        public List<DocumentDataDto> ReportITDDtos { get; set; }
    }
}

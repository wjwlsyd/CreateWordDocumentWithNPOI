using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    public class DocumentDataDto
    {
        public int UnderwritingYear { get; set; }
        public int TreatyId { get; set; }
        public string TreatyNumber { get; set; }
        public string TreatyName { get; set; }
        public FeeTypeEnum FeeType { get; set; }
        public decimal BusinessInCAD { get; set; }
        public decimal BusinessInUSD { get; set; }
    }
}

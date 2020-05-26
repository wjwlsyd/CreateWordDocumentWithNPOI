using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    public static class TestData
    {
        public static List<DocumentDataSet> CreateDocumentDataSets()
        {
            var dataSets = new List<DocumentDataSet>();

            var dataSet1 = new DocumentDataSet()
            {
                WorkSheetType = WorkSheetTypeEnum.Summary,
                WorkSheetName = "SUMMARY",
                UnderwritingYear = 2020,
                AssumingCompanyName = "KOR",
                QuarterEndDate = DateTime.Parse("2020-03-31"),
                ReportDtos = new List<DocumentDataDto>()
                 {
                     new DocumentDataDto()
                     {
                        UnderwritingYear = 2020,
                        BusinessInCAD = 29999,
                        BusinessInUSD = 399999,
                        FeeType = FeeTypeEnum.TotalManagementFees,
                        TreatyId = 1,
                        TreatyName = "TREATY-00001",
                        TreatyNumber = "TN-A00001"
                     },
                      new DocumentDataDto()
                     {
                        UnderwritingYear = 2020,
                        BusinessInCAD = 500,
                        BusinessInUSD = 100,
                        FeeType = FeeTypeEnum.TotalProfitCommisionManagementFeeAdjustments,
                        TreatyId = 1,
                        TreatyName = "TREATY-00001",
                        TreatyNumber = "TN-A00001"
                     },
                       new DocumentDataDto()
                     {
                        UnderwritingYear = 2020,
                        BusinessInCAD = 8888,
                        BusinessInUSD = 9999,
                        FeeType = FeeTypeEnum.TotalManagementFees,
                        TreatyId = 2,
                        TreatyName = "TREATY-00002",
                        TreatyNumber = "TN-A00002"
                     },
                        new DocumentDataDto()
                     {
                        UnderwritingYear = 2020,
                        BusinessInCAD = 7777,
                        BusinessInUSD = 6666,
                        FeeType = FeeTypeEnum.TotalProfitCommisionManagementFeeAdjustments,
                        TreatyId = 2,
                        TreatyName = "TREATY-00002",
                        TreatyNumber = "TN-A00002"
                     },
                      new DocumentDataDto()
                     {
                        UnderwritingYear = 2021,
                        BusinessInCAD = 678,
                        BusinessInUSD = 789,
                        FeeType = FeeTypeEnum.TotalManagementFees,
                        TreatyId = 1,
                        TreatyName = "TREATY-00001",
                        TreatyNumber = "TN-A00001"
                     }
                }
            };
            dataSets.Add(dataSet1);
            return dataSets;
        }
    }
}

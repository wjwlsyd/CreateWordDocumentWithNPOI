using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordDocument
{
    public class GlobalConstants
    {
        public static string GenerateFolderPathAppSettingName = "GenerateFolderPath";

        private static string _CurrentBaseFolder = AppDomain.CurrentDomain.BaseDirectory;

        public static string CurrentBaseFolder
        {
            get { return _CurrentBaseFolder; }
            set { _CurrentBaseFolder = value; }
        }

        public static string GenerateFolderPath = CurrentBaseFolder + GlobalHelper.GetAppConfig(GenerateFolderPathAppSettingName) +"\\";
    }

    public enum FeeTypeEnum
    {
        TotalManagementFees = 1,
        TotalProfitCommisionManagementFeeAdjustments
    }
    public enum WorkSheetTypeEnum
    {
        Summary,
        UnderwritingYear,
        ITD
    }
}

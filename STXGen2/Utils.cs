using SAPbobsCOM;
using System;

namespace STXGen2
{
    internal class Utils
    {
        public static SAPbobsCOM.Company oCompany;

        public static string decSep { get; private set; }
        public static string thousSep { get; private set; }
        public static int MeasureDec { get; private set; }
        public static int PriceDec { get; private set; }
        public static int SumDec { get; private set; }
        public static int QtyDec { get; private set; }
        public static string MainCurrency { get; private set; }
        public static string SystemCurrency { get; private set; }
        public static string DirectRate { get; private set; }

        internal static void CompSettings()
        {
            string sSql = $"select \"DecSep\",\"ThousSep\",\"MeasureDec\",\"PriceDec\",\"SumDec\",\"QtyDec\",\"MainCurncy\",\"SysCurrncy\",\"DirectRate\" from \"OADM\"";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);
            if (!rs.EoF)
            {
                decSep = string.IsNullOrEmpty((string)rs.Fields.Item("DecSep").Value) ? "" : (string)rs.Fields.Item("DecSep").Value;
                thousSep = string.IsNullOrEmpty((string)rs.Fields.Item("ThousSep").Value) ? "" : (string)rs.Fields.Item("ThousSep").Value;
                MeasureDec = (int)rs.Fields.Item("MeasureDec").Value;
                PriceDec = (int)rs.Fields.Item("PriceDec").Value;
                SumDec = (int)rs.Fields.Item("SumDec").Value;
                QtyDec = (int)rs.Fields.Item("QtyDec").Value;
                MainCurrency = (string)rs.Fields.Item("MainCurncy").Value;
                SystemCurrency = (string)rs.Fields.Item("SysCurrncy").Value;
                DirectRate = (string)rs.Fields.Item("DirectRate").Value;
            }
        }
    }
}
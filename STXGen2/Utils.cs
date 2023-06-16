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

        public static System.Globalization.NumberFormatInfo GetSAPNumberFormatInfo()
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = new System.Globalization.NumberFormatInfo();
            sapNumberFormat.NumberDecimalSeparator = Utils.decSep;
            sapNumberFormat.NumberGroupSeparator = Utils.thousSep;
            return sapNumberFormat;
        }

        public static System.Globalization.CultureInfo GetCompanyCulture()
        {
            SAPbouiCOM.BoLanguages language = Program.SBO_Application.Language;
            string cultureCode = "en-US";

            switch (Program.SBO_Application.Language)
            {
                case SAPbouiCOM.BoLanguages.ln_Hebrew:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Spanish_Ar:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Polish:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Spanish_Pa:
                    break;
                case SAPbouiCOM.BoLanguages.ln_English:
                case SAPbouiCOM.BoLanguages.ln_English_Gb:
                case SAPbouiCOM.BoLanguages.ln_English_Sg:
                    System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = new System.Globalization.CultureInfo("en-GB");
                    break;
                case SAPbouiCOM.BoLanguages.ln_German:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Serbian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Danish:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Norwegian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Italian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Hungarian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Chinese:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Dutch:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Finnish:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Greek:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Portuguese:
                    System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = new System.Globalization.CultureInfo("pt-PT");
                    break;
                case SAPbouiCOM.BoLanguages.ln_Swedish:
                    break;
                case SAPbouiCOM.BoLanguages.ln_English_Cy:
                    break;
                case SAPbouiCOM.BoLanguages.ln_French:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Spanish:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Russian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Spanish_La:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Czech_Cz:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Slovak_Sk:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Korean_Kr:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Portuguese_Br:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Japanese_Jp:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Turkish_Tr:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Ukrainian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_TrdtnlChinese_Hk:
                    break;
                default:
                    break;
            }
            return new System.Globalization.CultureInfo(cultureCode);
        }
    }
}